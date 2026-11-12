import type { DynamicDateUpdateFormat } from "../types/field.js";

const formats: readonly DynamicDateUpdateFormat[] = ["application-default", "month-day-year", "weekday-month-day-year", "day-month-year", "month-day-year-long", "automatic-month-day-year", "day-short-month-year", "month-year", "short-month-year", "month-day-year-time-12", "month-day-year-time-12-seconds", "time-24", "time-24-seconds", "time-12", "time-12-seconds"];
export function isDynamicDateUpdateFormat(value: unknown): value is DynamicDateUpdateFormat { return typeof value === "string" && formats.includes(value as DynamicDateUpdateFormat); }
