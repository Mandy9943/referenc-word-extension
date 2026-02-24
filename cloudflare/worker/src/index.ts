import { Container } from "@cloudflare/containers";

export class WorkflowContainer extends Container {
  defaultPort = 8080;
  sleepAfter = "20m";
  enableInternet = true;
}

type WorkflowContainerStub = {
  startAndWaitForPorts: () => Promise<void>;
  fetch: (request: Request) => Promise<Response>;
};

type WorkflowContainerNamespace = {
  getByName: (name: string) => WorkflowContainerStub;
};

type Env = {
  WORKFLOW_CONTAINER: WorkflowContainerNamespace;
  WORKFLOW_CONTAINER_NAME?: string;
};

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "content-type": "application/json; charset=utf-8" },
  });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    const { pathname } = new URL(request.url);

    if (pathname === "/worker-health") {
      return jsonResponse({ ok: true, service: "referenc-workflow-cloud" });
    }

    const containerName = env.WORKFLOW_CONTAINER_NAME || "workflow-web-singleton";
    const container = env.WORKFLOW_CONTAINER.getByName(containerName);

    await container.startAndWaitForPorts();
    return container.fetch(request);
  },
};
