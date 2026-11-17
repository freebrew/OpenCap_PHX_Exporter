import { cpSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const source = join(root, "native-host");
const packed = join(root, "dist", "native-host");
const stage = join(packed, "stage");
rmSync(stage, { recursive: true, force: true });
mkdirSync(stage, { recursive: true });
for (const name of [
  "phx-outlook-host.ps1",
  "phx-outlook-host.cmd",
  "install-host.ps1",
  "install-host.cmd",
  "repair-host.cmd",
  "uninstall-host.cmd",
  "host-version.json",
  "extension-id.txt",
  "README.txt",
]) {
  cpSync(join(source, name), join(stage, name));
}
const zip = join(packed, "phx-outlook-host.zip");
rmSync(zip, { force: true });
const result = spawnSync("powershell.exe", [
  "-NoProfile",
  "-Command",
  `Compress-Archive -Path '${stage}\\*' -DestinationPath '${zip}' -Force`,
], { stdio: "inherit" });
if (result.status !== 0) process.exit(result.status ?? 1);
rmSync(stage, { recursive: true, force: true });
console.log(zip);
