import { spawnSync } from "node:child_process";

const workflow = "deploy-presentation-preview.yml";
const ref = "main";

function run(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit" });
  if (result.error) {
    throw result.error;
  }
  if (result.status !== 0) {
    const error = new Error(`${command} ${args.join(" ")} exited with status ${result.status ?? 1}.`);
    error.exitCode = result.status ?? 1;
    throw error;
  }
}

try {
  run("gh", ["--version"]);
  try {
    run("gh", ["auth", "status"]);
  } catch (error) {
    console.error("GitHub CLI is not logged in. Please run: gh auth login");
    process.exit(error.exitCode ?? 1);
  }
  run("gh", ["workflow", "run", workflow, "--ref", ref]);
  console.log(`Triggered GitHub Actions: ${workflow} (${ref})`);
  console.log("Check run status: https://github.com/openHacking/pptkit/actions");
  console.log("Deploy URL: https://openhacking.github.io/pptkit/");
} catch (error) {
  console.error("Unable to trigger GitHub Pages deployment.");
  if (error.code === "ENOENT") {
    console.error("Please install GitHub CLI (gh) and run: gh auth login");
  } else {
    console.error(error.message);
  }
  process.exit(error.exitCode ?? 1);
}
