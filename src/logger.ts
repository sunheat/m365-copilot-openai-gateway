import type { GatewayLogFormat, GatewayLogLevel } from "./config.js";

const LOG_LEVEL_RANK: Record<Exclude<GatewayLogLevel, "silent">, number> = {
  error: 1,
  warn: 2,
  info: 3,
  debug: 4,
  trace: 5,
};

export type LogFields = Record<string, boolean | number | string | null | undefined>;
export type LogWriter = (level: Exclude<GatewayLogLevel, "silent">, line: string) => void;

export interface GatewayLogger {
  error(event: string, fields?: LogFields): void;
  warn(event: string, fields?: LogFields): void;
  info(event: string, fields?: LogFields): void;
  debug(event: string, fields?: LogFields): void;
  trace(event: string, fields?: LogFields): void;
}

function defaultWriter(level: Exclude<GatewayLogLevel, "silent">, line: string): void {
  const stream = level === "error" || level === "warn" ? process.stderr : process.stdout;
  stream.write(`${line}\n`);
}

function cleanFields(fields: LogFields): Record<string, boolean | number | string | null> {
  return Object.fromEntries(
    Object.entries(fields).filter((entry): entry is [string, boolean | number | string | null] => entry[1] !== undefined),
  );
}

function prettyValue(value: boolean | number | string | null): string {
  return typeof value === "string" ? JSON.stringify(value) : String(value);
}

export function createGatewayLogger(
  configuredLevel: GatewayLogLevel,
  format: GatewayLogFormat,
  writer: LogWriter = defaultWriter,
): GatewayLogger {
  const write = (level: Exclude<GatewayLogLevel, "silent">, event: string, fields: LogFields): void => {
    if (configuredLevel === "silent" || LOG_LEVEL_RANK[level] > LOG_LEVEL_RANK[configuredLevel]) return;

    const record = {
      timestamp: new Date().toISOString(),
      level,
      event,
      ...cleanFields(fields),
    };
    if (format === "json") {
      writer(level, JSON.stringify(record));
      return;
    }

    const fieldText = Object.entries(record)
      .filter(([key]) => key !== "timestamp" && key !== "level" && key !== "event")
      .map(([key, value]) => `${key}=${prettyValue(value)}`)
      .join(" ");
    writer(level, `${record.timestamp} ${level.toUpperCase()} ${event}${fieldText ? ` ${fieldText}` : ""}`);
  };

  return {
    error: (event, fields = {}) => write("error", event, fields),
    warn: (event, fields = {}) => write("warn", event, fields),
    info: (event, fields = {}) => write("info", event, fields),
    debug: (event, fields = {}) => write("debug", event, fields),
    trace: (event, fields = {}) => write("trace", event, fields),
  };
}
