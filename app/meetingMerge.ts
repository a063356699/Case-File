export type MeetingRecord = { id: string; propertyNo?: string; caseName?: string; address?: string; lastModifiedAt?: string; archived?: string; status?: string };

// Prefer the latest formal record when a property occurs in more than one week.
// Never use the address alone: separate contracts can share an address.
export function uniqueMeetingRecords<T extends MeetingRecord>(records: T[]): T[] {
  const result = new Map<string, T>();
  for (const record of records) {
    const key = record.propertyNo?.trim().toUpperCase() || record.id;
    const previous = result.get(key);
    if (!previous || (Date.parse(record.lastModifiedAt || '') || 0) > (Date.parse(previous.lastModifiedAt || '') || 0)) result.set(key, record);
  }
  return [...result.values()];
}

export function isMeetingChange(record: MeetingRecord): boolean {
  return !!record.archived || /售出|租出|成交下架|下架/.test(record.status || '');
}
