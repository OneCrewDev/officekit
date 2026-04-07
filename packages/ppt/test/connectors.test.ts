import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  addConnector,
  getConnectors,
  setConnectorEndpoints,
  removeConnector,
  setConnectorStyle,
} from "../src/connectors.js";
import { addSlide } from "../src/slides.js";
import { addShape } from "../src/shapes.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-connectors-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("addConnector - adds a straight connector between two shapes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // First add two shapes to connect
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await addConnector(tempPath, 1, "straight", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "addConnector should succeed");
    assert.ok(result.data?.path, "Result should have path");
    assert.ok(result.data?.path.includes("/connector["), "Path should include connector index");
  } finally {
    // Clean up
  }
});

test("addConnector - adds an elbow connector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 3000000 }, { width: 2000000, height: 1500000 });

    const result = await addConnector(tempPath, 1, "elbow", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    assert.ok(result.ok, "addConnector with elbow type should succeed");
    assert.ok(result.data?.path.includes("/connector["), "Path should include connector index");
  } finally {
    // Clean up
  }
});

test("addConnector - adds a curved connector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 3000000 }, { width: 2000000, height: 1500000 });

    const result = await addConnector(tempPath, 1, "curved", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    assert.ok(result.ok, "addConnector with curved type should succeed");
  } finally {
    // Clean up
  }
});

test("addConnector - adds an arrow connector with options", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await addConnector(tempPath, 1, "arrow", "/slide[1]/shape[1]", "/slide[1]/shape[2]", {
      color: "FF0000",
      width: 25400,
      endArrow: "arrow",
      label: "Flow",
    });

    assert.ok(result.ok, "addConnector with arrow type and options should succeed");
  } finally {
    // Clean up
  }
});

test("addConnector - returns error for invalid connector type", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await addConnector(tempPath, 1, "invalid" as any, "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    assert.ok(!result.ok, "addConnector should fail for invalid connector type");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getConnectors - lists connectors on a slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    // Add a connector first
    await addConnector(tempPath, 1, "straight", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    const result = await getConnectors(tempPath, 1);

    assert.ok(result.ok, "getConnectors should succeed");
    assert.ok(result.data?.connectors, "Result should have connectors array");
    assert.ok(Array.isArray(result.data?.connectors), "Connectors should be an array");
  } finally {
    // Clean up
  }
});

test("getConnectors - returns empty array when no connectors", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getConnectors(tempPath, 1);

    assert.ok(result.ok, "getConnectors should succeed");
    assert.ok(Array.isArray(result.data?.connectors), "Connectors should be an array");
    assert.equal(result.data?.connectors.length, 0, "Should have no connectors");
  } finally {
    // Clean up
  }
});

test("getConnectors - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getConnectors(tempPath, 999);

    assert.ok(!result.ok, "getConnectors should fail for invalid slide index");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setConnectorEndpoints - updates connector endpoints", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add three shapes
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 7000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    // Add connector between shape 1 and 2
    const addResult = await addConnector(tempPath, 1, "straight", "/slide[1]/shape[1]", "/slide[1]/shape[2]");
    assert.ok(addResult.ok, "addConnector should succeed");

    // Update to connect shape 1 and 3
    const result = await setConnectorEndpoints(
      tempPath,
      "/slide[1]/connector[1]",
      "/slide[1]/shape[1]",
      "/slide[1]/shape[3]"
    );

    assert.ok(result.ok, "setConnectorEndpoints should succeed");
  } finally {
    // Clean up
  }
});

test("setConnectorEndpoints - returns error for invalid connector path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setConnectorEndpoints(tempPath, "invalid", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    assert.ok(!result.ok, "setConnectorEndpoints should fail for invalid path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeConnector - removes a connector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    // Add a connector
    const addResult = await addConnector(tempPath, 1, "straight", "/slide[1]/shape[1]", "/slide[1]/shape[2]");
    assert.ok(addResult.ok, "addConnector should succeed");

    // Get connectors before removal
    const beforeResult = await getConnectors(tempPath, 1);
    const countBefore = beforeResult.data?.connectors.length ?? 0;

    // Remove the connector
    const result = await removeConnector(tempPath, "/slide[1]/connector[1]");

    assert.ok(result.ok, "removeConnector should succeed");

    // Get connectors after removal
    const afterResult = await getConnectors(tempPath, 1);
    const countAfter = afterResult.data?.connectors.length ?? 0;

    assert.equal(countAfter, countBefore - 1, "Connector count should decrease by 1");
  } finally {
    // Clean up
  }
});

test("removeConnector - returns error for invalid connector path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeConnector(tempPath, "invalid");

    assert.ok(!result.ok, "removeConnector should fail for invalid path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeConnector - returns error for non-existent connector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeConnector(tempPath, "/slide[1]/connector[999]");

    assert.ok(!result.ok, "removeConnector should fail for non-existent connector");
    assert.equal(result.error?.code, "not_found");
  } finally {
    // Clean up
  }
});

test("setConnectorStyle - updates connector style", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    // Add a connector
    await addConnector(tempPath, 1, "straight", "/slide[1]/shape[1]", "/slide[1]/shape[2]");

    // Update the style
    const result = await setConnectorStyle(tempPath, "/slide[1]/connector[1]", {
      color: "00FF00",
      width: 25400,
      endArrow: "arrow",
    });

    assert.ok(result.ok, "setConnectorStyle should succeed");
  } finally {
    // Clean up
  }
});

test("setConnectorStyle - returns error for invalid connector path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setConnectorStyle(tempPath, "invalid", { color: "FF0000" });

    assert.ok(!result.ok, "setConnectorStyle should fail for invalid path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setConnectorStyle - returns error for non-existent connector", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setConnectorStyle(tempPath, "/slide[1]/connector[999]", { color: "FF0000" });

    assert.ok(!result.ok, "setConnectorStyle should fail for non-existent connector");
    assert.equal(result.error?.code, "not_found");
  } finally {
    // Clean up
  }
});
