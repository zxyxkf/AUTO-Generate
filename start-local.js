import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';

const rootDir = path.resolve(import.meta.dirname);
const port = Number(process.env.PORT || 5174);
const url = `http://127.0.0.1:${port}`;
const apiUrl = `${url}/api/templates`;

process.chdir(rootDir);

main().catch((error) => {
  console.error('');
  console.error('Startup failed:');
  console.error(error.message || error);
  console.error('');
  console.error('Keep this window open and send logs/start.log to the maintainer.');
  process.exitCode = 1;
});

async function main() {
  console.log('Office document generator');
  console.log(`Working directory: ${rootDir}`);
  console.log('');

  ensureNodeModules();
  ensureDist();

  if (await isHttpReady(apiUrl)) {
    console.log(`Local service is already running: ${url}`);
    openBrowser(url);
    return;
  }

  console.log('Starting local service...');
  const server = spawn(process.execPath, ['server/index.js'], {
    cwd: rootDir,
    env: { ...process.env, PORT: String(port) },
    stdio: ['ignore', 'pipe', 'pipe'],
  });

  server.stdout.on('data', (chunk) => process.stdout.write(chunk));
  server.stderr.on('data', (chunk) => process.stderr.write(chunk));
  server.on('exit', (code) => {
    console.log('');
    console.log(`Local service exited. Exit code: ${code ?? 'unknown'}`);
    process.exitCode = code || 0;
  });

  await waitForHttp(apiUrl, 15000);
  console.log('');
  console.log(`Started: ${url}`);
  console.log('The browser should open automatically. Keep this window open while using the tool.');
  openBrowser(url);

  process.on('SIGINT', () => {
    console.log('');
    console.log('Stopping local service...');
    server.kill();
    process.exit(0);
  });
}

function ensureNodeModules() {
  if (fs.existsSync(path.join(rootDir, 'node_modules', 'express'))) return;
  throw new Error('Missing node_modules. The portable package is incomplete.');
}

function ensureDist() {
  if (fs.existsSync(path.join(rootDir, 'dist', 'index.html'))) return;
  console.log('Missing dist/index.html. Building frontend...');
  const npmCmd = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const result = spawnSync(npmCmd, ['run', 'build'], {
    cwd: rootDir,
    stdio: 'inherit',
  });
  if (result.status !== 0) {
    throw new Error('Frontend build failed. Check npm run build output.');
  }
}

function openBrowser(targetUrl) {
  const command = process.platform === 'win32'
    ? ['cmd', ['/c', 'start', '', targetUrl]]
    : process.platform === 'darwin'
      ? ['open', [targetUrl]]
      : ['xdg-open', [targetUrl]];
  spawn(command[0], command[1], {
    detached: true,
    stdio: 'ignore',
  }).unref();
}

function isHttpReady(targetUrl) {
  return new Promise((resolve) => {
    const req = http.get(targetUrl, (res) => {
      res.resume();
      resolve(res.statusCode >= 200 && res.statusCode < 500);
    });
    req.on('error', () => resolve(false));
    req.setTimeout(1000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

async function waitForHttp(targetUrl, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (await isHttpReady(targetUrl)) return;
    await sleep(300);
  }
  throw new Error(`Local service startup timed out: ${targetUrl}`);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
