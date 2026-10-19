// API de teste: GET /table?rows=N[&delay=ms] devolve um array JSON de objetos planos (10 colunas).
// delay simula a latência de uma API real.
import http from "node:http";

const PORT = Number(process.env.PORT ?? 8765);
const cache = new Map();

function buildTableJson(rowCount) {
  const rows = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    rows[i] = {
      id: i + 1,
      codigo: `C${String(i).padStart(7, "0")}`,
      cliente: `Cliente ${i % 997}`,
      cidade: ["Sao Paulo", "Rio", "Curitiba", "Recife", "Manaus"][i % 5],
      data: new Date(Date.UTC(2026, 0, 1) + (i % 365) * 86400000).toISOString().slice(0, 10),
      quantidade: (i * 7) % 1000,
      preco: Math.round(((i * 13) % 10000) * 1.37) / 100,
      desconto: ((i * 3) % 50) / 100,
      ativo: i % 3 !== 0,
      categoria: `CAT-${i % 40}`,
    };
  }
  return JSON.stringify(rows);
}

http
  .createServer((req, res) => {
    const url = new URL(req.url, `http://localhost:${PORT}`);
    if (url.pathname !== "/table") {
      res.writeHead(404).end();
      return;
    }
    const rowCount = Number(url.searchParams.get("rows") ?? 1000);
    if (!cache.has(rowCount)) cache.set(rowCount, buildTableJson(rowCount));
    const body = cache.get(rowCount);
    const delayMs = Number(url.searchParams.get("delay") ?? 0);
    setTimeout(() => {
      res.writeHead(200, { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(body) });
      res.end(body);
    }, delayMs);
  })
  .listen(PORT, () => console.log(`listening on ${PORT}`));
