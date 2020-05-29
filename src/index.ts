import { SpreadSheetService } from "./spreadsheet.service";
import { DateService } from "./date.service";
import { TogglService } from "./toggl.service";
import CONFIG from "./config.service";

import { DetailReportItem } from "./toggl.service";

////////////////////////////////////////////////////////////////
// Batch
////////////////////////////////////////////////////////////////

// every 1 Hour
function sync() {
  // setup sheet for log
  const ss = new SpreadSheetService(
    PropertiesService.getScriptProperties().getProperty(
      "GOOGLE_DRIVE_LOG_DIR_ID"
    ),
    PropertiesService.getScriptProperties().getProperty(
      "SPREADSHEET_TEMPLATE_ID"
    ),
    CONFIG.SpreadSheet.FILE_NAME,
    CONFIG.SpreadSheet.SHEET_NAME,
    CONFIG.SpreadSheet.DATA_AREA_NUM.ROW,
    CONFIG.SpreadSheet.DATA_AREA_NUM.COLUM
  );

  // init toggl service
  const toggl = new TogglService(
    PropertiesService.getScriptProperties().getProperty("TOGGL_API_KEY"),
    PropertiesService.getScriptProperties().getProperty("TOGGL_USER_AGENT"),
    PropertiesService.getScriptProperties().getProperty("TOGGL_WORKSPACE_ID")
  );

  // get weekly reports of toggle
  const now = new Date();
  const startDate = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() - 1
  );
  const togglItems = toggl.getDetailReport(
    Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd")
  );
  if (!togglItems) {
    throw new Error("Failed to get toggl reports. Empty data in response.");
  }
  console.log(`SUCCEEDED: Fetch toggl data: ${togglItems.length} items`);

  insertItems(ss, togglItems);
  console.log("SUCCEEDED: Created/Updated toggl items");
}

////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////
// Private Functions
////////////////////////////////////////////////////////////////

function insertItems(ss: SpreadSheetService, items: Array<DetailReportItem>) {
  let counter = {
    create: 0,
    update: 0,
    skip: 0,
    delete: 0,
  };

  for (let item of items) {
    // check to duplicate id
    const rowNum = ss.getRowNumByValue(
      item.id.toString(),
      CONFIG.SpreadSheet.COLUMN.ID
    );

    // Get the position inserted the items
    let startRowNum: number = ss.getSheet().getLastRow() + 1;
    let startColumnNum: number = CONFIG.SpreadSheet.DATA_AREA_NUM.COLUM;

    // Update last-update-time value of item if item already exits
    if (rowNum >= CONFIG.SpreadSheet.DATA_AREA_NUM.ROW) {
      // compare last-update-time
      const oldLastUpdateTime = new Date(
        ss.getValue(rowNum, CONFIG.SpreadSheet.COLUMN.LAST_UPDATE)
      );
      const newLastUpdateTime = new Date(item.updated);
      if (oldLastUpdateTime.getTime() >= newLastUpdateTime.getTime()) {
        counter.skip += 1;
        continue;
      }

      // if new one > old one, overwrite item value
      startRowNum = rowNum;
      counter.update += 1;
    } else {
      counter.create += 1;
    }

    // create array included the elem of schedule item (object->array)
    const values = [
      [
        item.id,
        item.description,
        Utilities.formatDate(
          DateService.str2date(item.start),
          "Asia/Tokyo",
          SpreadSheetService.DATE_FORMAT
        ),
        Utilities.formatDate(
          DateService.str2date(item.end),
          "Asia/Tokyo",
          SpreadSheetService.DATE_FORMAT
        ),
        item.project,
        item.tags.join(","),
        "toggl",
        Utilities.formatDate(
          DateService.str2date(item.updated),
          "Asia/Tokyo",
          SpreadSheetService.DATE_FORMAT
        ),
      ],
    ];
    ss.setValues(values, startRowNum, startColumnNum, 1, values[0].length);
  }

  ////////////////////////////////////////////////////
  // TODO Delete schedule-items on SpreadSheet
  ////////////////////////////////////////////////////

  ////////////////////////////////////////////////////
  // Clean up
  ////////////////////////////////////////////////////
  console.log(
    `create:${counter.create}, update:${counter.update}, skip:${counter.skip}, delete:${counter.delete}`
  );

  // sort by start date
  if (counter.create > 0 || counter.update > 0) {
    ss.sortSheet(
      CONFIG.SpreadSheet.COLUMN.START,
      false,
      CONFIG.SpreadSheet.DATA_AREA_NUM.ROW,
      CONFIG.SpreadSheet.DATA_AREA_NUM.COLUM
    );
    console.log(`Detected to change table and sorted table by 'start-date'`);
  }
}
