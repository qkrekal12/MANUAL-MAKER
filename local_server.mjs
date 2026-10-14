import { createServer } from "node:http";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { execFile, spawn } from "node:child_process";
import { tmpdir } from "node:os";
import worker from "./site/worker/index.js";

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
const host = "127.0.0.1";
let port = Number(process.env.MANUAL_STUDIO_PORT || 8766);
const bundledModules = path.resolve(path.dirname(process.execPath), "..", "node_modules");
const localDocxScript = path.join(moduleDir, "node_modules", "docx", "dist", "index.iife.js");
const docxScript = existsSync(localDocxScript)
  ? localDocxScript
  : path.join(bundledModules, "docx", "dist", "index.iife.js");
const outlookScript = path.join(moduleDir, "outlook_mail.ps1");
const proposalHtml = path.join(moduleDir, "AX_PoC_기획서 (1).html");

function loadApiKey() {
  if (process.env.OPENAI_API_KEY) return process.env.OPENAI_API_KEY.trim();
  const envPaths = [path.join(moduleDir, ".env"), path.resolve(moduleDir, "..", ".env")];
  for (const envPath of envPaths) {
    if (!existsSync(envPath)) continue;
    const contents = readFileSync(envPath, "utf8").replace(/^\uFEFF/, "");
    for (const line of contents.split(/\r?\n/)) {
      const match = line.match(/^\s*(?:export\s+)?OPENAI_API_KEY\s*=\s*(.*?)\s*$/);
      if (!match) continue;
      let value = match[1];
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) value = value.slice(1, -1);
      if (value.trim()) return value.trim();
    }
  }
  return "";
}

const apiKey = loadApiKey();
function openBrowser(address) {
  return new Promise((resolve) => {
    execFile("rundll32.exe", ["url.dll,FileProtocolHandler", address], { windowsHide: true }, (error) => {
      if (!error) return resolve(true);
      execFile(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-Command", `Start-Process -FilePath '${address}'`],
        { windowsHide: true },
        (fallbackError) => resolve(!fallbackError),
      );
    });
  });
}

const send = (response, outgoing) => {
  outgoing.statusCode = response.status;
  for (const [name, value] of response.headers) outgoing.setHeader(name, value);
  response.arrayBuffer().then((body) => outgoing.end(Buffer.from(body))).catch(() => outgoing.end());
};

function runOutlook(action, payload) {
  return new Promise((resolve, reject) => {
    const child = spawn("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-STA", "-File", outlookScript, "-Action", action], {
      windowsHide: true,
      stdio: ["pipe", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", () => reject(new Error("Outlook을 시작하지 못했습니다. 클래식 Outlook이 설치되어 있는지 확인해 주세요.")));
    child.on("close", (code) => {
      if (code !== 0) return reject(new Error(stderr.trim() || "Outlook에서 이메일을 처리하지 못했습니다."));
      try {
        resolve(JSON.parse(stdout.trim()));
      } catch {
        reject(new Error("Outlook에서 예상하지 못한 응답을 받았습니다."));
      }
    });
    if (payload) child.stdin.end(JSON.stringify(payload));
    else child.stdin.end();
  });
}

async function readJsonBody(incoming, maxSize = 40 * 1024 * 1024) {
  const chunks = [];
  let size = 0;
  for await (const chunk of incoming) {
    size += chunk.length;
    if (size > maxSize) throw new Error("요청 파일이 너무 큽니다. 이미지 크기를 줄이고 다시 시도해 주세요.");
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

function exportPdfFromHtml(html) {
  return new Promise(async (resolve, reject) => {
    let tempDir;
    let child;
    const finish = async (error, pdf) => {
      clearTimeout(timer);
      if (tempDir) await rm(tempDir, { recursive: true, force: true }).catch(() => {});
      if (error) reject(error); else resolve(pdf);
    };
    const timer = setTimeout(() => {
      child?.kill();
      void finish(new Error("PDF 변환 시간이 초과되었습니다. 다시 시도해 주세요."));
    }, 90000);
    try {
      const edgeCandidates = [
        path.join(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)", "Microsoft", "Edge", "Application", "msedge.exe"),
        path.join(process.env.ProgramFiles || "C:\\Program Files", "Microsoft", "Edge", "Application", "msedge.exe"),
      ];
      const edge = edgeCandidates.find(existsSync);
      if (!edge) throw new Error("PDF 저장을 위해 Microsoft Edge가 필요합니다.");
      tempDir = await mkdtemp(path.join(tmpdir(), "manual-studio-pdf-"));
      const htmlPath = path.join(tempDir, "manual.html");
      const pdfPath = path.join(tempDir, "manual.pdf");
      await writeFile(htmlPath, html, "utf8");
      child = spawn(edge, [
        "--headless=new", "--no-sandbox", "--disable-gpu", "--disable-gpu-sandbox", "--disable-gpu-compositing", "--disable-software-rasterizer", "--disable-extensions", "--disable-features=VizDisplayCompositor", "--no-first-run",
        `--user-data-dir=${path.join(tempDir, "profile")}`,
        `--print-to-pdf=${pdfPath}`, "--no-pdf-header-footer", pathToFileURL(htmlPath).href,
      ], { windowsHide: true, stdio: ["ignore", "ignore", "pipe"] });
      let stderr = "";
      child.stderr.setEncoding("utf8");
      child.stderr.on("data", (chunk) => { stderr += chunk; });
      child.on("error", (error) => { void finish(new Error(`PDF 변환 프로그램을 시작하지 못했습니다: ${error.message}`)); });
      child.on("close", async (code) => {
        try {
          if (code !== 0 || !existsSync(pdfPath)) throw new Error(stderr.trim() || "Microsoft Edge가 PDF 파일을 만들지 못했습니다.");
          const pdf = await readFile(pdfPath);
          if (!pdf.subarray(0, 5).equals(Buffer.from("%PDF-"))) throw new Error("PDF 파일 형식을 확인하지 못했습니다.");
          await finish(null, pdf);
        } catch (error) { await finish(error); }
      });
    } catch (error) { await finish(error); }
  });
}

function jsonResponse(outgoing, status, payload) {
  outgoing.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store", "x-content-type-options": "nosniff" });
  outgoing.end(JSON.stringify(payload));
}

const server = createServer(async (incoming, outgoing) => {
  try {
    const url = new URL(incoming.url || "/", `http://${host}:${port}`);
    if (incoming.method === "GET" && url.pathname === "/docx.js") {
      if (!existsSync(docxScript)) {
        outgoing.writeHead(503, { "content-type": "text/plain; charset=utf-8" });
        outgoing.end("The bundled Word document library was not found.");
        return;
      }
      outgoing.writeHead(200, { "content-type": "text/javascript; charset=utf-8", "cache-control": "no-store", "x-content-type-options": "nosniff" });
      outgoing.end(await readFile(docxScript));
      return;
    }

    if (incoming.method === "GET" && url.pathname === "/") {
      if (!existsSync(proposalHtml)) {
        outgoing.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
        outgoing.end("기획서 첫 페이지 파일을 찾지 못했습니다.");
        return;
      }
      outgoing.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store", "x-content-type-options": "nosniff" });
      outgoing.end(await readFile(proposalHtml));
      return;
    }

    if (incoming.method === "GET" && url.pathname === "/studio") {
      const response = await worker.fetch(new Request(`http://${host}:${port}/`), { OPENAI_API_KEY: apiKey });
      const html = (await response.text()).replace(
        "https://cdn.jsdelivr.net/npm/docx@9.11.0/dist/index.iife.js",
        "/docx.js",
      );
      outgoing.writeHead(response.status, Object.fromEntries(response.headers));
      outgoing.end(html);
      return;
    }

    if (incoming.method === "GET" && url.pathname === "/api/mail/accounts") {
      try {
        const result = await runOutlook("accounts");
        jsonResponse(outgoing, 200, result);
      } catch (error) {
        jsonResponse(outgoing, 500, { error: error.message || "Outlook 계정을 확인하지 못했습니다." });
      }
      return;
    }

    if (incoming.method === "POST" && url.pathname === "/api/export/pdf") {
      const expectedOrigin = "http://" + host + ":" + server.address().port;
      if (incoming.headers.origin && incoming.headers.origin !== expectedOrigin) {
        jsonResponse(outgoing, 403, { error: "허용되지 않은 요청입니다. 매뉴얼 스튜디오 페이지에서 다시 시도해 주세요." });
        return;
      }
      let payload;
      try {
        payload = await readJsonBody(incoming, 45 * 1024 * 1024);
      } catch (error) {
        jsonResponse(outgoing, 400, { error: error.message || "PDF 변환 요청을 읽지 못했습니다." });
        return;
      }
      try {
        if (typeof payload.html !== "string" || payload.html.length > 40 * 1024 * 1024 || !/<html[\s>]/i.test(payload.html)) {
          throw new Error("PDF로 저장할 매뉴얼 내용이 올바르지 않습니다.");
        }
        const pdf = await exportPdfFromHtml(payload.html);
        const baseName = path.basename(String(payload.filename || "사용자_매뉴얼.pdf")).replace(/\.pdf$/i, "");
        const filename = encodeURIComponent(baseName + ".pdf");
        outgoing.writeHead(200, {
          "content-type": "application/pdf",
          "content-disposition": "attachment; filename=\"manual.pdf\"; filename*=UTF-8''" + filename,
          "content-length": pdf.length,
          "cache-control": "no-store",
          "x-content-type-options": "nosniff",
        });
        outgoing.end(pdf);
      } catch (error) {
        jsonResponse(outgoing, 500, { error: error.message || "PDF를 변환하지 못했습니다." });
      }
      return;
    }

    if (incoming.method === "POST" && url.pathname === "/api/mail/send") {
      const expectedOrigin = "http://" + host + ":" + server.address().port;
      if (incoming.headers.origin && incoming.headers.origin !== expectedOrigin) {
        jsonResponse(outgoing, 403, { error: "허용되지 않은 요청입니다. 매뉴얼 스튜디오 페이지에서 다시 시도해 주세요." });
        return;
      }
      let payload;
      try {
        payload = await readJsonBody(incoming);
      } catch (error) {
        jsonResponse(outgoing, 400, { error: error.message || "이메일 요청을 읽지 못했습니다." });
        return;
      }
      try {
        const result = await runOutlook("send", payload);
        jsonResponse(outgoing, 200, result);
      } catch (error) {
        jsonResponse(outgoing, 500, { error: error.message || "Outlook 발송에 실패했습니다." });
      }
      return;
    }

    if (url.pathname === "/api/generate" && incoming.method === "POST") {
      const chunks = [];
      let size = 0;
      for await (const chunk of incoming) {
        size += chunk.length;
        if (size > 36 * 1024 * 1024) {
          outgoing.writeHead(413, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
          outgoing.end(JSON.stringify({ error: "이미지 전체 크기는 30MB 이하로 줄여 주세요." }));
          incoming.destroy();
          return;
        }
        chunks.push(chunk);
      }
      const headers = new Headers();
      if (incoming.headers["content-type"]) headers.set("content-type", incoming.headers["content-type"]);
      headers.set("content-length", String(size));
      const body = Buffer.concat(chunks);
      const request = new Request(`http://${host}:${port}/api/generate`, { method: "POST", headers, body });
      send(await worker.fetch(request, { OPENAI_API_KEY: apiKey }), outgoing);
      return;
    }

    outgoing.writeHead(404, { "content-type": "text/plain; charset=utf-8" });
    outgoing.end("Not found");
  } catch {
    if (!outgoing.headersSent) outgoing.writeHead(500, { "content-type": "text/plain; charset=utf-8" });
    outgoing.end("로컬 웹 앱에서 요청을 처리하지 못했습니다.");
  }
});

server.on("error", async (error) => {
  if (error.code === "EADDRINUSE") {
    const existingAddress = `http://${host}:${port}`;
    try {
      const response = await fetch(existingAddress, { signal: AbortSignal.timeout(1200) });
      const page = await response.text();
      if (response.ok && (page.includes("매뉴얼 스튜디오") || page.includes("AI 사용자 매뉴얼"))) {
        const opened = await openBrowser(existingAddress);
        console.log(opened
          ? "매뉴얼 스튜디오가 이미 실행 중이어서 해당 페이지를 브라우저에서 열었습니다."
          : `브라우저를 자동으로 열지 못했습니다. 다음 주소를 복사해 여세요: ${existingAddress}`);
        return;
      }
    } catch {}

    console.log(`포트 ${port}가 다른 프로그램에서 사용 중입니다. 사용 가능한 포트를 찾습니다.`);
    port = 0;
    server.listen(port, host);
    return;
  }
  console.error("웹 서버를 시작하지 못했습니다.");
  process.exitCode = 1;
});

server.listen(port, host, () => {
  const actualPort = server.address().port;
  const address = `http://${host}:${actualPort}`;
  console.log(`매뉴얼 스튜디오가 실행 중입니다: ${address}`);
  console.log(apiKey ? "API 키: 로컬 .env에서 안전하게 읽었습니다." : "API 키: 설정되지 않았습니다. 프로그램 폴더의 .env를 확인해 주세요.");
  console.log("이미지와 API 키는 이 PC에서 브라우저 서버가 처리하며 이미지 파일은 저장하지 않습니다.");
  if (process.env.MANUAL_STUDIO_NO_OPEN === "1") return;
  openBrowser(address).then((opened) => {
    if (!opened) console.log(`브라우저를 자동으로 열지 못했습니다. 다음 주소를 복사해 여세요: ${address}`);
  });
});
