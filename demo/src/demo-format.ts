const useAnsi = Boolean(process.stdout.isTTY && !process.env.NO_COLOR);

function wrap(code: string, text: string): string {
  return useAnsi ? `\x1b[${code}m${text}\x1b[0m` : text;
}

export const fmt = {
  policy(header: string, lines: string[] = []): void {
    console.log(wrap("36;1", `[POLICY] ${header}`));
    for (const line of lines) {
      console.log(wrap("36", `  ${line}`));
    }
  },
  run(command: string): void {
    console.log(wrap("33;1", `[RUN] ${command}`));
  },
  log(line: string): void {
    console.log(`  ${line}`);
  },
  osErr(line: string): void {
    console.log(wrap("31", `  stderr (OS): ${line}`));
  },
  pass(summary: string): void {
    console.log(wrap("32;1", `✓ PASS — ${summary}`));
  },
  blocked(summary: string): void {
    console.log(wrap("31;1", `✕ BLOCKED — ${summary}`));
  },
  why(explanation: string): void {
    console.log(wrap("2", `Why: ${explanation}`));
  },
};
