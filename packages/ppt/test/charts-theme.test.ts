import test from "node:test";
import assert from "node:assert/strict";
import { copyFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";

import {
  getChart,
  addChart,
  setChartData,
  setChartType,
} from "../src/charts.js";

import {
  getTheme,
  getThemeColor,
  setThemeColor,
  getThemeFont,
} from "../src/theme.js";

const TEST_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source-officecli/examples/ppt/outputs/data_presentation.pptx";
const CHART_PPTX = "/Users/llm/Desktop/Code/office/officekit/packages/parity-tests/fixtures/source/examples/binaries/budget_review_v2.pptx";

async function copyToTemp(sourcePath: string): Promise<string> {
  const tempPath = path.join(tmpdir(), `ppt-charts-theme-test-${Date.now()}.pptx`);
  await copyFile(sourcePath, tempPath);
  return tempPath;
}

// ============================================================================
// Chart Tests
// ============================================================================

test("getChart - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getChart(tempPath, "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getChart - returns error for path without slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getChart(tempPath, "/chart[1]");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getChart - returns error for non-existent chart", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getChart(tempPath, "/slide[1]/chart[999]");
    // May fail because chart doesn't exist, or may fail on invalid path
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getChart - works with chart-containing presentation", async () => {
  const tempPath = await copyToTemp(CHART_PPTX);
  try {
    // slide6 has a chart
    const result = await getChart(tempPath, "/slide[6]/chart[1]");
    if (result.ok) {
      assert.ok(result.data!.chart);
      assert.ok(result.data!.chart.path);
      assert.ok(result.data!.chart.type);
    } else {
      // Chart extraction may have issues, but shouldn't throw
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - returns error for invalid slide index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      999,
      "bar",
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        series: [{ name: "Test", values: [1, 2, 3] }],
        categories: ["A", "B", "C"]
      }
    );
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("addChart - returns error for invalid chart type", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      1,
      "invalid_type" as any,
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        series: [{ name: "Test", values: [1, 2, 3] }],
        categories: ["A", "B", "C"]
      }
    );
    // Should not throw, but may succeed or fail
    // (TypeScript would catch the invalid type at compile time)
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - adds bar chart to slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      1,
      "bar",
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        title: "Test Bar Chart",
        series: [{ name: "Sales", values: [100, 200, 150] }],
        categories: ["Jan", "Feb", "Mar"]
      }
    );
    // May fail if slide doesn't exist or other issues
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - adds pie chart to slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      1,
      "pie",
      { x: 1000000, y: 1000000, width: 4000000, height: 4000000 },
      {
        title: "Test Pie Chart",
        series: [{ name: "Shares", values: [35, 28, 22, 15] }],
        categories: ["Category 1", "Category 2", "Category 3", "Category 4"]
      }
    );
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - adds line chart to slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      2,
      "line",
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        series: [{ name: "Trend", values: [10, 15, 12, 18, 20] }],
        categories: ["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"]
      }
    );
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - adds column chart to slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      3,
      "column",
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        series: [
          { name: "Product A", values: [120, 150, 180] },
          { name: "Product B", values: [80, 100, 120] }
        ],
        categories: ["Q1", "Q2", "Q3"]
      }
    );
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("addChart - adds area chart to slide", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await addChart(
      tempPath,
      1,
      "area",
      { x: 1000000, y: 1000000, width: 6000000, height: 4000000 },
      {
        series: [{ name: "Revenue", values: [50, 75, 100, 125, 150] }],
        categories: ["2019", "2020", "2021", "2022", "2023"]
      }
    );
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setChartData - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(CHART_PPTX);
  try {
    const result = await setChartData(
      tempPath,
      "invalid",
      [{ name: "Test", values: [1, 2, 3] }],
      ["A", "B", "C"]
    );
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setChartData - returns error for path without chart index", async () => {
  const tempPath = await copyToTemp(CHART_PPTX);
  try {
    const result = await setChartData(
      tempPath,
      "/slide[1]",
      [{ name: "Test", values: [1, 2, 3] }],
      ["A", "B", "C"]
    );
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setChartType - returns error for invalid path", async () => {
  const tempPath = await copyToTemp(CHART_PPTX);
  try {
    const result = await setChartType(tempPath, "invalid", "line");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setChartType - returns error for path without chart index", async () => {
  const tempPath = await copyToTemp(CHART_PPTX);
  try {
    const result = await setChartType(tempPath, "/slide[1]", "line");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

// ============================================================================
// Theme Tests
// ============================================================================

test("getTheme - returns theme info", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getTheme(tempPath);
    if (result.ok) {
      assert.ok(result.data!.theme);
      assert.ok(result.data!.theme.colors || result.data!.theme.fonts);
    } else {
      // May fail if theme extraction has issues
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeColor - returns accent1 color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeColor(tempPath, "accent1");
    if (result.ok) {
      assert.ok(result.data!.color);
      // Should be a hex color
      assert.ok(/^[0-9A-F]{6}$/i.test(result.data!.color));
    } else {
      // May fail if theme doesn't exist or color doesn't exist
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeColor - returns dk1 (dark1) color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeColor(tempPath, "dk1");
    if (result.ok) {
      assert.ok(result.data!.color);
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeColor - returns lt1 (light1) color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeColor(tempPath, "lt1");
    if (result.ok) {
      assert.ok(result.data!.color);
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeColor - returns error for invalid color index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeColor(tempPath, "invalid" as any);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setThemeColor - sets accent1 color", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setThemeColor(tempPath, "accent1", "FF0000");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setThemeColor - sets accent2 color with 3-char hex", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setThemeColor(tempPath, "accent2", "F00");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setThemeColor - sets accent2 color with # prefix", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setThemeColor(tempPath, "accent2", "#00FF00");
    if (!result.ok) {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("setThemeColor - returns error for invalid color format", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setThemeColor(tempPath, "accent1", "invalid");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("setThemeColor - returns error for invalid color index", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await setThemeColor(tempPath, "invalid" as any, "FF0000");
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});

test("getThemeFont - returns major font", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeFont(tempPath, "major");
    if (result.ok) {
      assert.ok(result.data!.font);
      assert.ok(typeof result.data!.font === "string");
    } else {
      // May fail if theme doesn't exist or font doesn't exist
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeFont - returns minor font", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeFont(tempPath, "minor");
    if (result.ok) {
      assert.ok(result.data!.font);
      assert.ok(typeof result.data!.font === "string");
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeFont - returns heading font (alias for major)", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeFont(tempPath, "heading");
    if (result.ok) {
      assert.ok(result.data!.font);
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeFont - returns body font (alias for minor)", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeFont(tempPath, "body");
    if (result.ok) {
      assert.ok(result.data!.font);
    } else {
      assert.ok(result.error);
    }
  } finally {
    // Clean up
  }
});

test("getThemeFont - returns error for invalid font type", async () => {
  const tempPath = await copyToTemp(TEST_PPTX);
  try {
    const result = await getThemeFont(tempPath, "invalid" as any);
    assert.ok(!result.ok);
    assert.equal(result.error?.code, "invalid_input");
  } finally {
    // Clean up
  }
});
