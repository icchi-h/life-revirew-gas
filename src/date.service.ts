export class DateService {
  static str2date(iso8601: string) {
    return new Date(Date.parse(iso8601));
  }
}
