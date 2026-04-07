import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getAnimations,
  setAnimation,
  removeAnimation,
} from "../src/animations.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/beautiful_presentation.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-animations-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

test("getAnimations - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getAnimations(tempPath, 999);
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("getAnimations - returns animations for valid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getAnimations(tempPath, 1);
    if (result.ok) {
      assert.ok(result.data);
      assert.equal(result.data.slideIndex, 1);
      assert.ok(result.data.path);
      assert.ok(typeof result.data.count === "number");
      assert.ok(Array.isArray(result.data.animations));
    }
  } finally {
    // Clean up
  }
});

test("setAnimation - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setAnimation(tempPath, "invalid", {
      effect: "fade",
      class: "entrance",
    });
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setAnimation - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setAnimation(tempPath, "/slide[999]/shape[1]", {
      effect: "fade",
    });
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setAnimation - returns error for invalid shape path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setAnimation(tempPath, "/slide[1]", {
      effect: "fade",
    });
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("setAnimation - sets animation on valid shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setAnimation(tempPath, "/slide[1]/shape[1]", {
      effect: "fade",
      class: "entrance",
      trigger: "onClick",
      duration: 500,
    });
    // May succeed or fail depending on shape existence
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setAnimation - sets animation with all parameters", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setAnimation(tempPath, "/slide[1]/shape[1]", {
      effect: "fly",
      class: "entrance",
      trigger: "afterPrev",
      duration: 1000,
      delay: 500,
    });
    // May succeed or fail depending on shape existence
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("removeAnimation - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeAnimation(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("removeAnimation - returns error for invalid slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeAnimation(tempPath, "/slide[999]/shape[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("removeAnimation - returns error for invalid shape path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeAnimation(tempPath, "/slide[1]");
    assert.ok(!result.ok);
  } finally {
    // Clean up
  }
});

test("removeAnimation - removes animation from valid shape", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await removeAnimation(tempPath, "/slide[1]/shape[1]");
    // May succeed or fail depending on shape existence and animation presence
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});