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
  GEMINI_API_KEY?: string;
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
    const forwardHeaders = new Headers(request.headers);
    const geminiKey = (env.GEMINI_API_KEY || "").trim();

    if (geminiKey) {
      // Worker secret is injected only on the edge side, so forward it to container runtime.
      forwardHeaders.set("x-workflow-gemini-key", geminiKey);
    }

    const forwardedRequest = new Request(request, { headers: forwardHeaders });

    await container.startAndWaitForPorts();
    return container.fetch(forwardedRequest);
  },
};
