#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import https from "node:https";

// ---------------------------------------------------------------------------
// Aria Orchestrator REST API client
// ---------------------------------------------------------------------------

class AriaClient {
  constructor(baseUrl, username, password, ignoreCert = true) {
    const parsed = new URL(baseUrl);
    this.host = parsed.hostname;
    this.port = parsed.port || 443;
    this.basePath = parsed.pathname.replace(/\/+$/, "") || "/vco/api";
    this.auth = Buffer.from(`${username}:${password}`).toString("base64");
    this.ignoreCert = ignoreCert;
  }

  /** Generic REST request returning parsed JSON. */
  request(method, path, body = null) {
    return new Promise((resolve, reject) => {
      const opts = {
        hostname: this.host,
        port: this.port,
        path: `${this.basePath}${path}`,
        method,
        headers: {
          Accept: "application/json",
          Authorization: `Basic ${this.auth}`,
        },
        rejectUnauthorized: !this.ignoreCert,
      };

      if (body) {
        const payload = JSON.stringify(body);
        opts.headers["Content-Type"] = "application/json";
        opts.headers["Content-Length"] = Buffer.byteLength(payload);
      }

      const req = https.request(opts, (res) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => {
          const raw = Buffer.concat(chunks).toString();
          // 202 Accepted returns Location header for new executions
          if (res.statusCode === 202) {
            resolve({
              status: res.statusCode,
              location: res.headers.location,
              body: raw ? JSON.parse(raw) : null,
            });
            return;
          }
          if (res.statusCode === 204) {
            resolve({ status: 204, body: null });
            return;
          }
          if (res.statusCode >= 400) {
            reject(
              new Error(
                `HTTP ${res.statusCode} ${res.statusMessage}: ${raw}`
              )
            );
            return;
          }
          try {
            resolve(JSON.parse(raw));
          } catch {
            resolve(raw);
          }
        });
      });
      req.on("error", reject);
      if (body) req.write(JSON.stringify(body));
      req.end();
    });
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getClient() {
  const url = process.env.ARIA_URL;
  const user = process.env.ARIA_USERNAME;
  const pass = process.env.ARIA_PASSWORD;
  if (!url || !user || !pass) {
    throw new Error(
      "Missing environment variables: ARIA_URL, ARIA_USERNAME, ARIA_PASSWORD"
    );
  }
  const ignoreCert = process.env.ARIA_IGNORE_CERT !== "false";
  return new AriaClient(url, user, pass, ignoreCert);
}

function formatJson(obj) {
  return JSON.stringify(obj, null, 2);
}

// ---------------------------------------------------------------------------
// MCP Server
// ---------------------------------------------------------------------------

const server = new McpServer({
  name: "mcp-aria-orchestrator",
  version: "1.0.0",
});

// ---- Workflows ------------------------------------------------------------

server.tool(
  "list_workflows",
  "List all workflows in Aria Orchestrator. Optionally filter by name.",
  {
    filter: z
      .string()
      .optional()
      .describe("Optional name substring to filter workflows"),
  },
  async ({ filter }) => {
    const client = getClient();
    const qs = filter
      ? `?conditions=name~${encodeURIComponent(filter)}`
      : "";
    const result = await client.request("GET", `/workflows${qs}`);
    const workflows = (result.link || []).map((l) => ({
      id: l.attributes?.find((a) => a.name === "id")?.value,
      name: l.attributes?.find((a) => a.name === "name")?.value,
      description: l.attributes?.find((a) => a.name === "description")?.value,
      href: l.href,
    }));
    return {
      content: [
        {
          type: "text",
          text: `Found ${workflows.length} workflow(s):\n\n${formatJson(workflows)}`,
        },
      ],
    };
  }
);

server.tool(
  "get_workflow",
  "Get detailed information about a specific workflow including input/output parameters.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
  },
  async ({ workflowId }) => {
    const client = getClient();
    const result = await client.request("GET", `/workflows/${workflowId}`);
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "run_workflow",
  "Execute a workflow with the given input parameters. Returns the execution ID for status polling.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
    parameters: z
      .array(
        z.object({
          name: z.string().describe("Parameter name"),
          type: z
            .string()
            .describe("Parameter type (string, number, boolean, etc.)"),
          value: z.string().describe("Parameter value as string"),
        })
      )
      .optional()
      .describe("Input parameters for the workflow"),
  },
  async ({ workflowId, parameters }) => {
    const client = getClient();
    const body = { parameters: [] };
    if (parameters) {
      body.parameters = parameters.map((p) => ({
        type: p.type,
        name: p.name,
        scope: "local",
        value: {
          [p.type]: { value: p.value },
        },
      }));
    }
    const result = await client.request(
      "POST",
      `/workflows/${workflowId}/executions`,
      body
    );
    const execId = result.location
      ? result.location.split("/").pop()
      : null;
    return {
      content: [
        {
          type: "text",
          text: `Workflow execution started.\nExecution ID: ${execId}\nStatus: ${result.status}\nLocation: ${result.location}`,
        },
      ],
    };
  }
);

server.tool(
  "get_execution",
  "Get the status and details of a workflow execution.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
    executionId: z.string().describe("The execution ID (UUID)"),
  },
  async ({ workflowId, executionId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/workflows/${workflowId}/executions/${executionId}`
    );
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "get_execution_state",
  "Get just the current state of a workflow execution (running, completed, error, etc.).",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
    executionId: z.string().describe("The execution ID (UUID)"),
  },
  async ({ workflowId, executionId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/workflows/${workflowId}/executions/${executionId}/state`
    );
    return {
      content: [{ type: "text", text: `Execution state: ${formatJson(result)}` }],
    };
  }
);

server.tool(
  "get_execution_logs",
  "Get logs from a workflow execution.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
    executionId: z.string().describe("The execution ID (UUID)"),
  },
  async ({ workflowId, executionId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/workflows/${workflowId}/executions/${executionId}/logs`
    );
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "cancel_execution",
  "Cancel a running workflow execution.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
    executionId: z.string().describe("The execution ID (UUID)"),
  },
  async ({ workflowId, executionId }) => {
    const client = getClient();
    await client.request(
      "DELETE",
      `/workflows/${workflowId}/executions/${executionId}`
    );
    return {
      content: [{ type: "text", text: `Execution ${executionId} cancelled.` }],
    };
  }
);

server.tool(
  "list_executions",
  "List all executions for a workflow.",
  {
    workflowId: z.string().describe("The workflow ID (UUID)"),
  },
  async ({ workflowId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/workflows/${workflowId}/executions`
    );
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

// ---- Actions ---------------------------------------------------------------

server.tool(
  "list_actions",
  "List all actions (scriptable tasks) in Aria Orchestrator. Optionally filter by name.",
  {
    filter: z
      .string()
      .optional()
      .describe("Optional name substring to filter actions"),
  },
  async ({ filter }) => {
    const client = getClient();
    const qs = filter
      ? `?conditions=name~${encodeURIComponent(filter)}`
      : "";
    const result = await client.request("GET", `/actions${qs}`);
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "get_action",
  "Get detailed information about a specific action including its script content.",
  {
    actionId: z.string().describe("The action ID (UUID)"),
  },
  async ({ actionId }) => {
    const client = getClient();
    const result = await client.request("GET", `/actions/${actionId}`);
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

// ---- Configuration Elements ------------------------------------------------

server.tool(
  "list_configurations",
  "List all configuration elements in Aria Orchestrator.",
  {},
  async () => {
    const client = getClient();
    const result = await client.request("GET", "/configurations");
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "get_configuration",
  "Get a specific configuration element and its attributes.",
  {
    configId: z.string().describe("The configuration element ID (UUID)"),
  },
  async ({ configId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/configurations/${configId}`
    );
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

// ---- Resource Elements -----------------------------------------------------

server.tool(
  "list_resources",
  "List all resource elements in Aria Orchestrator.",
  {},
  async () => {
    const client = getClient();
    const result = await client.request("GET", "/resources");
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

server.tool(
  "get_resource",
  "Get a specific resource element.",
  {
    resourceId: z.string().describe("The resource element ID (UUID)"),
  },
  async ({ resourceId }) => {
    const client = getClient();
    const result = await client.request(
      "GET",
      `/resources/${resourceId}`
    );
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

// ---- Packages --------------------------------------------------------------

server.tool(
  "list_packages",
  "List all packages in Aria Orchestrator.",
  {},
  async () => {
    const client = getClient();
    const result = await client.request("GET", "/packages");
    return {
      content: [{ type: "text", text: formatJson(result) }],
    };
  }
);

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error("Fatal:", err);
  process.exit(1);
});
