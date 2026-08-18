import { describe, expect, it } from "vitest";
import { createGatewayLogger } from "../src/logger.js";

describe("createGatewayLogger", () => {
  it("filters records below the configured verbosity", () => {
    const lines: string[] = [];
    const logger = createGatewayLogger("info", "pretty", (_level, line) => lines.push(line));

    logger.error("request_failed", { code: "gateway_upstream_error" });
    logger.info("request_completed", { status: 200 });
    logger.debug("stream_snapshot", { snapshot: 1 });

    expect(lines).toHaveLength(2);
    expect(lines[0]).toContain("ERROR request_failed");
    expect(lines[1]).toContain("INFO request_completed status=200");
  });

  it("supports structured JSON records and silent mode", () => {
    const lines: string[] = [];
    const logger = createGatewayLogger("silent", "json", (_level, line) => lines.push(line));
    logger.info("request_completed", { status: 200, omitted: undefined });
    expect(lines).toEqual([]);

    const jsonLines: string[] = [];
    const jsonLogger = createGatewayLogger("info", "json", (_level, line) => jsonLines.push(line));
    jsonLogger.info("request_completed", { status: 200, omitted: undefined });
    expect(JSON.parse(jsonLines[0] ?? "")).toMatchObject({
      level: "info",
      event: "request_completed",
      status: 200,
    });
    expect(jsonLines[0]).not.toContain("omitted");
  });
});
