const fs = require("fs");
const path = require("path");
const readline = require("readline");
const { spawnSync } = require("child_process");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const DEPLOYMENT_ID = "AKfycbxPnzc3avDB_b44FsjPY6jgubJzGgmpakuH7jXlbFZEG_mSfsKTWuMmBM6LZvwGnxLHBA";

const DEFAULT_SOURCES = {
  clasp: path.join(PROJECT_ROOT, "prepared", "clasp.json"),
  config: path.join(PROJECT_ROOT, "prepared", "Config.gs"),
};

function run(cmd, args) {
  console.log(`\n> ${cmd} ${args.join(" ")}`);
  const result = spawnSync(cmd, args, {
    cwd: PROJECT_ROOT,
    stdio: "inherit",
    shell: false,
  });

  if (result.error) {
    throw result.error;
  }

  if (result.status !== 0) {
    throw new Error(`${cmd} exited with code ${result.status}`);
  }
}

function ensureFile(filePath, label) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`${label} not found: ${filePath}`);
  }
}

function copyFile(sourcePath, targetPath, label) {
  ensureFile(sourcePath, label);
  fs.copyFileSync(sourcePath, targetPath);
  console.log(`Copied ${label}: ${sourcePath} -> ${targetPath}`);
}

function ask(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

async function main() {
  const sourceClasp = process.env.SIMOD_SOURCE_CLASP || DEFAULT_SOURCES.clasp;
  const sourceConfig = process.env.SIMOD_SOURCE_CONFIG || DEFAULT_SOURCES.config;
  const targetClasp = path.join(PROJECT_ROOT, ".clasp.json");
  const targetConfig = path.join(PROJECT_ROOT, "src", "Config.gs");

  console.log("SIMOD HPS deploy automation");
  console.log(`Project root: ${PROJECT_ROOT}`);
  console.log(`Deployment ID: ${DEPLOYMENT_ID}`);

  const updateNote = await ask("Update note: ");
  if (!updateNote) {
    throw new Error("Update note is required.");
  }

  copyFile(sourceClasp, targetClasp, "clasp.json");
  copyFile(sourceConfig, targetConfig, "Config.gs");

  run("npm", ["install"]);
  run("npm", ["run", "push"]);
  run("npx", ["clasp", "deploy", "--deploymentId", DEPLOYMENT_ID, "--description", updateNote]);

  console.log("\nDeployment finished.");
}

main().catch((error) => {
  console.error(`\nDeploy failed: ${error.message}`);
  process.exit(1);
});
