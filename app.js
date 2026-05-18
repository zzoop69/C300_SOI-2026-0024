const http = require("http");

const PORT = process.env.PORT || 3000;

const page = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Simple Node.js Webpage</title>
  <style>
    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-height: 100vh;
      font-family: Arial, sans-serif;
      color: #18212f;
      background: #f4f7fb;
      display: grid;
      place-items: center;
      padding: 24px;
    }

    main {
      width: min(720px, 100%);
      background: #ffffff;
      border: 1px solid #d9e2ec;
      border-radius: 8px;
      padding: 40px;
      box-shadow: 0 18px 45px rgba(24, 33, 47, 0.12);
    }

    h1 {
      margin: 0 0 12px;
      font-size: 2.25rem;
      line-height: 1.1;
    }

    p {
      margin: 0;
      color: #52606d;
      font-size: 1.1rem;
      line-height: 1.6;
    }

    .badge {
      display: inline-block;
      margin-bottom: 18px;
      padding: 6px 10px;
      border-radius: 999px;
      background: #e0f2fe;
      color: #075985;
      font-size: 0.85rem;
      font-weight: 700;
    }
  </style>
</head>
<body>
  <main>
    <span class="badge">Node.js</span>
    <h1>Hello, welcome to my simple webpage.</h1>
    <p>This page is served from <strong>app.js</strong> using Node's built-in HTTP server.</p>
  </main>
</body>
</html>`;

const server = http.createServer((request, response) => {
  if (request.url === "/" || request.url === "/index.html") {
    response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    response.end(page);
    return;
  }

  response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
  response.end("Page not found");
});

server.listen(PORT, () => {
  console.log(`Server is running at http://localhost:${PORT}`);
});
