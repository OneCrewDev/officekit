export interface SessionRecord {
  kind: "watch" | "resident";
  filePath: string;
  pid: number;
  startedAt: string;
  url?: string;
  port?: number;
  format?: string;
  socketPath?: string;
}

export function getSessionFilePath(kind: "watch" | "resident", filePath: string): string;
export function writeSessionRecord(kind: "watch" | "resident", filePath: string, record: SessionRecord): Promise<string>;
export function readSessionRecord(kind: "watch" | "resident", filePath: string): Promise<SessionRecord | null>;
export function removeSessionRecord(kind: "watch" | "resident", filePath: string): Promise<void>;
export function waitForSessionRecord(kind: "watch" | "resident", filePath: string, timeoutMs?: number): Promise<SessionRecord | null>;
