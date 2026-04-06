import test from "node:test";
import assert from "node:assert/strict";

import {
  ok,
  okVoid,
  err,
  fail,
  failWith,
  isOk,
  isErr,
  someIf,
  fromPromise,
  fromFn,
  map,
  mapErr,
  andThen,
  getOrElse,
  unwrap,
  unwrapOr,
  ErrorCodes,
  invalidPath,
  notFound,
  alreadyExists,
  invalidInput,
  notSupported,
} from "../src/result.ts";

test("ok - creates successful result", () => {
  const result = ok({ key: "value" });
  assert.ok(result.ok);
  assert.deepEqual(result.data, { key: "value" });
  assert.ok(result.error === undefined);
});

test("okVoid - creates successful void result", () => {
  const result = okVoid();
  assert.ok(result.ok);
  assert.ok(result.data === undefined);
});

test("err - creates error result", () => {
  const result = err("not_found", "Slide not found");
  assert.ok(!result.ok);
  assert.ok(result.error);
  assert.equal(result.error.code, "not_found");
  assert.equal(result.error.message, "Slide not found");
});

test("err - creates error with suggestion", () => {
  const result = err("not_found", "Slide not found", "Check the slide index");
  assert.ok(!result.ok);
  assert.equal(result.error.suggestion, "Check the slide index");
});

test("fail - creates error from Error object", () => {
  const error = new Error("Something went wrong");
  const result = fail(error);
  assert.ok(!result.ok);
  assert.equal(result.error.code, "unknown_error");
  assert.equal(result.error.message, "Something went wrong");
});

test("fail - creates error with custom code", () => {
  const error = new Error("Custom error");
  const result = fail(error, "custom_code");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "custom_code");
});

test("failWith - creates error with full details", () => {
  const result = failWith("validation_error", "Invalid input", "Provide a valid number");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "validation_error");
  assert.equal(result.error.message, "Invalid input");
  assert.equal(result.error.suggestion, "Provide a valid number");
});

test("isOk - type guard for successful results", () => {
  const success = ok({ data: 123 });
  const failure = err("code", "msg");

  assert.ok(isOk(success));
  assert.ok(!isOk(failure));
});

test("isErr - type guard for failed results", () => {
  const success = ok({ data: 123 });
  const failure = err("code", "msg");

  assert.ok(!isErr(success));
  assert.ok(isErr(failure));
});

test("someIf - converts nullable to result", () => {
  const value = { found: true };
  const nullValue = null;
  const undefinedValue = undefined;

  const foundResult = someIf(value, "not_found", "Value not found");
  assert.ok(foundResult.ok);
  assert.deepEqual(foundResult.data, value);

  const nullResult = someIf(nullValue, "not_found", "Value not found");
  assert.ok(!nullResult.ok);
  assert.equal(nullResult.error.code, "not_found");

  const undefinedResult = someIf(undefinedValue, "not_found", "Value not found");
  assert.ok(!undefinedResult.ok);
});

test("fromPromise - wraps async function", async () => {
  const successPromise = Promise.resolve({ value: 42 });
  const failPromise = Promise.reject(new Error("Async error"));

  const successResult = await fromPromise(successPromise);
  assert.ok(successResult.ok);
  assert.deepEqual(successResult.data, { value: 42 });

  const failResult = await fromPromise(failPromise);
  assert.ok(!failResult.ok);
  assert.equal(failResult.error.message, "Async error");
});

test("fromFn - wraps sync function", () => {
  const successFn = () => ({ value: 42 });
  const failFn = () => {
    throw new Error("Sync error");
  };

  const successResult = fromFn(successFn);
  assert.ok(successResult.ok);
  assert.deepEqual(successResult.data, { value: 42 });

  const failResult = fromFn(failFn);
  assert.ok(!failResult.ok);
  assert.equal(failResult.error.message, "Sync error");
});

test("map - transforms successful result data", () => {
  const result = ok({ count: 5 });
  const mapped = map(result, (d) => ({ ...d, doubled: d.count * 2 }));

  assert.ok(mapped.ok);
  assert.deepEqual(mapped.data, { count: 5, doubled: 10 });
});

test("map - passes through error", () => {
  const result = err("code", "message");
  const mapped = map(result, (d) => ({ ...d, doubled: d }));

  assert.ok(!mapped.ok);
  assert.equal(mapped.error.code, "code");
});

test("mapErr - transforms error", () => {
  const result = err("old_code", "Something failed");
  const mapped = mapErr(result, (e) => ({
    ...e,
    code: "new_" + e.code,
  }));

  assert.ok(!mapped.ok);
  assert.equal(mapped.error.code, "new_old_code");
  assert.equal(mapped.error.message, "Something failed");
});

test("mapErr - passes through success", () => {
  const result = ok({ data: true });
  const mapped = mapErr(result, (e) => ({ ...e, code: "should_not_change" }));

  assert.ok(mapped.ok);
});

test("andThen - chains result-returning functions", () => {
  const initial = ok(5);
  const chainFn = (n: number) =>
    n > 0 ? ok(n * 2) : err("invalid", "Must be positive");

  const successChain = andThen(initial, chainFn);
  assert.ok(successChain.ok);
  assert.equal(successChain.data, 10);

  const failChain = andThen(err("initial", "Initial error"), chainFn);
  assert.ok(!failChain.ok);
  assert.equal(failChain.error.code, "initial");
});

test("getOrElse - returns fallback on error", () => {
  const errorResult = err("code", "message");
  const fallback = { default: true };

  assert.deepEqual(getOrElse(errorResult, fallback), fallback);
  assert.deepEqual(getOrElse(ok({ found: true }), fallback), { found: true });
});

test("unwrap - extracts data or throws", () => {
  const result = ok({ key: "value" });
  assert.deepEqual(unwrap(result), { key: "value" });

  const errorResult = err("code", "message");
  assert.throws(() => unwrap(errorResult), /message/);
});

test("unwrapOr - extracts data or returns default", () => {
  const result = ok({ found: true });
  const defaultValue = { found: false };

  assert.deepEqual(unwrapOr(result, defaultValue), { found: true });

  const errorResult = err("code", "message");
  assert.deepEqual(unwrapOr(errorResult, defaultValue), { found: false });
});

test("ErrorCodes - contains standard error codes", () => {
  assert.equal(ErrorCodes.INVALID_PATH, "invalid_path");
  assert.equal(ErrorCodes.NOT_FOUND, "not_found");
  assert.equal(ErrorCodes.ALREADY_EXISTS, "already_exists");
  assert.equal(ErrorCodes.NOT_SUPPORTED, "not_supported");
  assert.equal(ErrorCodes.INVALID_INPUT, "invalid_input");
  assert.equal(ErrorCodes.INVALID_FORMAT, "invalid_format");
  assert.equal(ErrorCodes.TIMEOUT, "timeout");
  assert.equal(ErrorCodes.PERMISSION_DENIED, "permission_denied");
  assert.equal(ErrorCodes.OPERATION_FAILED, "operation_failed");
  assert.equal(ErrorCodes.UNKNOWN_ERROR, "unknown_error");
  assert.equal(ErrorCodes.USAGE_ERROR, "usage_error");
});

test("invalidPath - creates INVALID_PATH error", () => {
  const result = invalidPath("Path must be absolute", "Start with /");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "invalid_path");
  assert.equal(result.error.message, "Path must be absolute");
  assert.equal(result.error.suggestion, "Start with /");
});

test("notFound - creates NOT_FOUND error", () => {
  const result = notFound("Slide", "5", "Check slide count");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "not_found");
  assert.equal(result.error.message, "Slide 5 not found");
  assert.equal(result.error.suggestion, "Check slide count");
});

test("alreadyExists - creates ALREADY_EXISTS error", () => {
  const result = alreadyExists("Slide", "3", "Use --force to replace");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "already_exists");
  assert.equal(result.error.message, "Slide 3 already exists");
});

test("invalidInput - creates INVALID_INPUT error", () => {
  const result = invalidInput("Index must be positive", "Use index >= 1");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "invalid_input");
  assert.equal(result.error.message, "Index must be positive");
});

test("notSupported - creates NOT_SUPPORTED error", () => {
  const result = notSupported("This operation is not yet implemented");
  assert.ok(!result.ok);
  assert.equal(result.error.code, "not_supported");
  assert.equal(result.error.message, "This operation is not yet implemented");
});
