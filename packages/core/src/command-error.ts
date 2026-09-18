/** Rejected command or state transition. */
export class CommandError extends Error {
  /** Stable error discriminator. */
  readonly tag = "CommandError";
  /** Explain the rejected operation; state remains unchanged. */
  constructor(message: string) {
    super(`Workbook command rejected: ${message}`);
  }
}
