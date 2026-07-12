import { spawn } from "node:child_process";

function run(label, command, args, env = {}) {
  const child = spawn(command, args, {
    cwd: process.cwd(),
    env: {
      ...process.env,
      ...env,
    },
    stdio: "inherit",
  });

  child.on("exit", (code, signal) => {
    if (signal) {
      shutdown(signal);
      return;
    }

    if (code !== 0) {
      process.exitCode = code ?? 1;
      console.error(`${label} exited with code ${code}`);
      shutdown("SIGTERM");
    }
  });

  child.on("error", (error) => {
    process.exitCode = 1;
    console.error(`${label} failed: ${String(error)}`);
    shutdown("SIGTERM");
  });

  return child;
}

const processes = [
  run("server", "pnpm", ["exec", "tsx", "watch", "src/server.ts"], { PORT: "3210" }),
  run("client", "pnpm", ["exec", "vite"]),
];

function shutdown(signal) {
  for (const child of processes) {
    if (!child.killed) {
      child.kill(signal);
    }
  }
}

process.on("SIGINT", () => shutdown("SIGINT"));
process.on("SIGTERM", () => shutdown("SIGTERM"));
