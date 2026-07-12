export interface CliRunResult {
  exitCode: number;
  output: string;
}

export function runCli(args: string[] = []): CliRunResult {
  if (args.includes("--help")) {
    return {
      exitCode: 0,
      output: "PPTKit CLI bootstrap. Commands will be added as packages mature.",
    };
  }

  return {
    exitCode: 0,
    output: "PPTKit CLI bootstrap. No commands are implemented yet.",
  };
}

