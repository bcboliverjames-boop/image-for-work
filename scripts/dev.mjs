import { getHttpsServerOptions } from "office-addin-dev-certs";
import { createServer } from "vite";

const https = await getHttpsServerOptions();

const server = await createServer({
  server: {
    host: "::",
    port: 3002,
    https
  }
});

await server.listen();
server.printUrls();
