// DB(SpreadSheet) Setting

enum Column {
  ID = 1,
  SUBJECT,
  START,
  END,
  PROJECT,
  TAG,
  SOURCE,
  LAST_UPDATE,
}

const CONFIG = {
  SpreadSheet: {
    FILE_NAME: "schedule_latest",
    LOGFILE_PREFIX: "schedule",
    SHEET_NAME: "log",
    // start: 1~
    DATA_AREA_NUM: {
      ROW: 2,
      COLUM: 1,
    },
    COLUMN: Column,
  },
};

export default CONFIG;
