import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  groupShapes,
  ungroupShapes,
  getGroupChildren,
  addShapeToGroup,
  removeShapeFromGroup,
  getGroup,
} from "../src/groups.js";
import { addShape } from "../src/shapes.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-groups-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("groupShapes - groups two shapes together", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes to group
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "groupShapes should succeed");
    assert.ok(result.data?.path, "Result should have path");
    assert.ok(result.data?.path.includes("/group["), "Path should include group index");
  } finally {
    // Clean up
  }
});

test("groupShapes - returns error for less than 2 shapes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]"]);

    assert.ok(!result.ok, "groupShapes should fail with less than 2 shapes");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("groupShapes - returns error for shapes on different slides", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add shapes on different slides
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 2, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const result = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[2]/shape[1]"]);

    assert.ok(!result.ok, "groupShapes should fail for shapes on different slides");
  } finally {
    // Clean up
  }
});

test("ungroupShapes - dissolves a group and keeps shapes", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Ungroup
    const result = await ungroupShapes(tempPath, "/slide[1]/group[1]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "ungroupShapes should succeed");
  } finally {
    // Clean up
  }
});

test("ungroupShapes - returns error for invalid group path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await ungroupShapes(tempPath, "/slide[1]/shape[1]");

    assert.ok(!result.ok, "ungroupShapes should fail for non-group path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("ungroupShapes - returns error for non-existent group", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await ungroupShapes(tempPath, "/slide[1]/group[999]");

    assert.ok(!result.ok, "ungroupShapes should fail for non-existent group");
    assert.equal(result.error?.code, "not_found");
  } finally {
    // Clean up
  }
});

test("getGroupChildren - lists children in a group", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Get group children
    const result = await getGroupChildren(tempPath, "/slide[1]/group[1]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "getGroupChildren should succeed");
    assert.ok(result.data?.children, "Result should have children array");
    assert.equal(result.data?.children.length, 2, "Group should have 2 children");
  } finally {
    // Clean up
  }
});

test("getGroupChildren - returns error for invalid group path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getGroupChildren(tempPath, "/slide[1]/shape[1]");

    assert.ok(!result.ok, "getGroupChildren should fail for non-group path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addShapeToGroup - adds a shape to an existing group", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add three shapes and group the first two
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 7000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Add third shape to the group
    const result = await addShapeToGroup(tempPath, "/slide[1]/group[1]", "/slide[1]/shape[3]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "addShapeToGroup should succeed");

    // Verify group now has 3 children
    const childrenResult = await getGroupChildren(tempPath, "/slide[1]/group[1]");
    assert.ok(childrenResult.ok, "getGroupChildren should succeed");
    assert.equal(childrenResult.data?.children.length, 3, "Group should now have 3 children");
  } finally {
    // Clean up
  }
});

test("addShapeToGroup - returns error for non-existent shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Try to add non-existent shape
    const result = await addShapeToGroup(tempPath, "/slide[1]/group[1]", "/slide[1]/shape[999]");

    assert.ok(!result.ok, "addShapeToGroup should fail for non-existent shape");
    assert.equal(result.error?.code, "not_found");
  } finally {
    // Clean up
  }
});

test("removeShapeFromGroup - removes a shape from a group", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add three shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 7000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]", "/slide[1]/shape[3]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Remove first shape from group
    const result = await removeShapeFromGroup(tempPath, "/slide[1]/group[1]", "/slide[1]/group[1]/shape[1]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "removeShapeFromGroup should succeed");

    // Verify group now has 2 children
    const childrenResult = await getGroupChildren(tempPath, "/slide[1]/group[1]");
    assert.ok(childrenResult.ok, "getGroupChildren should succeed");
    assert.equal(childrenResult.data?.children.length, 2, "Group should now have 2 children");
  } finally {
    // Clean up
  }
});

test("removeShapeFromGroup - returns error for invalid shape path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Try to remove with invalid shape path
    const result = await removeShapeFromGroup(tempPath, "/slide[1]/group[1]", "/slide[1]/shape[1]");

    assert.ok(!result.ok, "removeShapeFromGroup should fail for invalid shape path");
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getGroup - returns group information", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    // Add two shapes and group them
    await addShape(tempPath, 1, "rectangle", { x: 1000000, y: 1000000 }, { width: 2000000, height: 1500000 });
    await addShape(tempPath, 1, "rectangle", { x: 4000000, y: 1000000 }, { width: 2000000, height: 1500000 });

    const groupResult = await groupShapes(tempPath, "/slide[1]", ["/slide[1]/shape[1]", "/slide[1]/shape[2]"]);
    assert.ok(groupResult.ok, "groupShapes should succeed");

    // Get group info
    const result = await getGroup(tempPath, "/slide[1]/group[1]");

    if (!result.ok) {
      console.log("Error:", result.error);
    }
    assert.ok(result.ok, "getGroup should succeed");
    assert.ok(result.data?.group, "Result should have group");
    assert.ok(result.data?.group.path, "Group should have path");
    assert.ok(result.data?.group.childCount !== undefined, "Group should have childCount");
    assert.equal(result.data?.group.childCount, 2, "Group should have 2 children");
  } finally {
    // Clean up
  }
});

test("getGroup - returns error for non-existent group", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getGroup(tempPath, "/slide[1]/group[999]");

    assert.ok(!result.ok, "getGroup should fail for non-existent group");
    assert.equal(result.error?.code, "not_found");
  } finally {
    // Clean up
  }
});
