import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";

const port = Number(process.env.PORT || 4173);
const host = process.env.HOST || "127.0.0.1";
const root = process.cwd();

const server = createServer(async (request, response) => {
  const url = new URL(request.url || "/", `http://localhost:${port}`);

  if (url.pathname === "/") {
    response.writeHead(302, {
      Location: "/stock_holdings_percent_dashboard.html",
    });
    response.end();
    return;
  }

  const requested = decodeURIComponent(url.pathname).replace(/^\/+/, "");
  const filePath = path.join(root, "public", requested);

  if (!filePath.startsWith(path.join(root, "public"))) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  try {
    const body = await readFile(filePath);
    response.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
    });
    response.end(body);
  } catch {
    response.writeHead(404);
    response.end("Not found");
  }
});

server.listen(port, host, () => {
  console.log(`Stock holdings site preview: http://${host}:${port}`);
});
