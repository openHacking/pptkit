import { spawn } from "node:child_process";

function run(cmd, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      stdio: "inherit",
      ...options,
    });

    child.on("exit", (code, signal) => {
      if (signal) {
        reject(new Error(`${cmd} exited with signal ${signal}`));
        return;
      }

      if (code !== 0) {
        reject(new Error(`${cmd} exited with code ${code}`));
        return;
      }

      resolve();
    });

    child.on("error", reject);
  });
}

await run("pnpm", ["--filter", "@pptkit/dev-app", "dev"]);
