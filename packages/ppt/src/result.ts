/**
 * Result envelope helpers for @officekit/ppt.
 *
 * All public API functions return a standardized Result<T> envelope:
 *   { ok: boolean, data?: T, error?: { code: string, message: string } }
 *
 * This module provides utilities for creating, composing, and handling
 * result envelopes in a type-safe manner.
 */

import type { Result, ResultError } from "./types.js";

// ============================================================================
// Result Constructors
// ============================================================================

/**
 * Creates a successful result.
 *
 * @example
 * ok({ slides: [slide1, slide2] })
 * // Returns: { ok: true, data: { slides: [slide1, slide2] } }
 */
export function ok<T>(data: T): Result<T> {
  return { ok: true, data };
}

/**
 * Creates a successful result with undefined data (for void returns).
 *
 * @example
 * ok()  // Returns: { ok: true }
 */
export function okVoid(): Result<void> {
  return { ok: true };
}

/**
 * Creates a failed result.
 *
 * @example
 * err("not_found", "Slide 5 does not exist")
 * // Returns: { ok: false, error: { code: "not_found", message: "Slide 5 does not exist" } }
 */
export function err(code: string, message: string, suggestion?: string): Result<never> {
  return {
    ok: false,
    error: suggestion ? { code, message, suggestion } : { code, message },
  };
}

/**
 * Creates a failed result from an error object.
 *
 * @example
 * fail(new Error("Something went wrong"))
 * // Returns: { ok: false, error: { code: "unknown_error", message: "Something went wrong" } }
 */
export function fail(error: Error, code = "unknown_error"): Result<never> {
  return { ok: false, error: { code, message: error.message } };
}

/**
 * Creates a failed result from an OfficekitError-like object.
 *
 * @example
 * failWith("not_found", "Slide not found", "Check that the slide index is within range")
 */
export function failWith(code: string, message: string, suggestion?: string): Result<never> {
  return err(code, message, suggestion);
}

// ============================================================================
// Result Helpers
// ============================================================================

/**
 * Type guard to check if a result is successful.
 *
 * @example
 * if (isOk(result)) {
 *   console.log(result.data);  // TypeScript knows result.data exists here
 * }
 */
export function isOk<T>(result: Result<T>): result is Result<T> & { data: T } {
  return result.ok === true;
}

/**
 * Type guard to check if a result is a failure.
 *
 * @example
 * if (isErr(result)) {
 *   console.error(result.error);  // TypeScript knows result.error exists here
 * }
 */
export function isErr<T>(result: Result<T>): result is Result<T> & { error: ResultError } {
  return result.ok === false;
}

/**
 * Converts a nullable value to a Result.
 * If value is null/undefined, returns a failure with the given error.
 * Otherwise returns success with the value.
 *
 * @example
 * const result = someIf(slide, "not_found", "Slide not found");
 */
export function someIf<T>(
  value: T | null | undefined,
  errorCode: string,
  errorMessage: string,
  suggestion?: string,
): Result<T> {
  if (value == null) {
    return err(errorCode, errorMessage, suggestion);
  }
  return ok(value);
}

/**
 * Wraps a promise, returning a Result instead of throwing.
 *
 * @example
 * const result = await fromPromise(fetchSlide(1));
 * if (isOk(result)) {
 *   console.log(result.data);
 * }
 */
export async function fromPromise<T>(
  promise: Promise<T>,
  errorCode = "operation_failed",
): Promise<Result<T>> {
  try {
    const data = await promise;
    return ok(data);
  } catch (e) {
    if (e instanceof Error) {
      return err(errorCode, e.message);
    }
    return err(errorCode, String(e));
  }
}

/**
 * Wraps a synchronous function, returning a Result instead of throwing.
 *
 * @example
 * const result = fromFn(() => parseSlideIndex("/slide[5]"));
 * if (isErr(result)) {
 *   console.error(result.error.code);
 * }
 */
export function fromFn<T>(fn: () => T, errorCode = "operation_failed"): Result<T> {
  try {
    return ok(fn());
  } catch (e) {
    if (e instanceof Error) {
      return err(errorCode, e.message);
    }
    return err(errorCode, String(e));
  }
}

// ============================================================================
// Result Combinators
// ============================================================================

/**
 * Maps the data inside a result if successful.
 *
 * @example
 * const result = ok({ count: 5 });
 * const mapped = map(result, (d) => ({ ...d, doubled: d.count * 2 }));
 * // Returns: { ok: true, data: { count: 5, doubled: 10 } }
 */
export function map<T, U>(result: Result<T>, fn: (data: T) => U): Result<U> {
  if (isErr(result)) {
    return result as Result<U>;
  }
  return ok(fn(result.data as T));
}

/**
 * Maps the error inside a result if failed.
 *
 * @example
 * const result = err("old_code", "Something failed");
 * const mapped = mapErr(result, (e) => ({
 *   ...e,
 *   code: "new_" + e.code
 * }));
 * // Returns: { ok: false, error: { code: "new_old_code", message: "Something failed" } }
 */
export function mapErr<T>(result: Result<T>, fn: (error: ResultError) => ResultError): Result<T> {
  if (!isErr(result)) {
    return result;
  }
  return { ok: false, error: fn(result.error) };
}

/**
 * Chains a result through a function that returns a result.
 *
 * @example
 * const result = ok(5);
 * const chained = andThen(result, (n) => n > 0 ? ok(n * 2) : err("invalid", "Must be positive"));
 */
export function andThen<T, U>(
  result: Result<T>,
  fn: (data: T) => Result<U>,
): Result<U> {
  if (isErr(result)) {
    return result as Result<U>;
  }
  return fn(result.data as T);
}

/**
 * Provides a fallback value if the result is a failure.
 *
 * @example
 * const result = err("not_found", "Not found");
 * const value = getOrElse(result, { default: true });
 * // Returns: { default: true }
 */
export function getOrElse<T>(result: Result<T>, fallback: T): T {
  if (isErr(result)) {
    return fallback;
  }
  return result.data as T;
}

/**
 * Unwraps the data from a result, throwing if it's a failure.
 * Use this only when you're certain the result is successful.
 *
 * @example
 * const value = unwrap(ok({ key: "value" }));
 * // Returns: { key: "value" }
 *
 * @throws Error if result is a failure
 */
export function unwrap<T>(result: Result<T>): T {
  if (isErr(result)) {
    throw new Error(result.error.message);
  }
  return result.data as T;
}

/**
 * Unwraps the data from a result, returning a default value if it's a failure.
 *
 * @example
 * const value = unwrapOr(ok({ key: "value" }), { default: true });
 * // Returns: { key: "value" }
 *
 * const value2 = unwrapOr(err("oops", "Failed"), { default: true });
 * // Returns: { default: true }
 */
export function unwrapOr<T>(result: Result<T>, defaultValue: T): T {
  if (isErr(result)) {
    return defaultValue;
  }
  return result.data as T;
}

// ============================================================================
// Common Error Codes
// ============================================================================

/**
 * Common error codes used throughout the PPT package.
 * These provide consistent machine-readable error codes for common failure cases.
 */
export const ErrorCodes = {
  /** Path format is invalid */
  INVALID_PATH: "invalid_path",
  /** Path references an element that doesn't exist */
  NOT_FOUND: "not_found",
  /** Path references an element that already exists (for add operations) */
  ALREADY_EXISTS: "already_exists",
  /** Operation is not supported */
  NOT_SUPPORTED: "not_supported",
  /** Operation failed due to invalid input */
  INVALID_INPUT: "invalid_input",
  /** File format is invalid or corrupted */
  INVALID_FORMAT: "invalid_format",
  /** Operation timed out */
  TIMEOUT: "timeout",
  /** Permission denied */
  PERMISSION_DENIED: "permission_denied",
  /** General operation failure */
  OPERATION_FAILED: "operation_failed",
  /** Unknown error */
  UNKNOWN_ERROR: "unknown_error",
  /** Usage error (invalid command syntax) */
  USAGE_ERROR: "usage_error",
} as const;

export type ErrorCode = (typeof ErrorCodes)[keyof typeof ErrorCodes];

// ============================================================================
// Error Factories
// ============================================================================

/**
 * Creates an INVALID_PATH error.
 *
 * @example
 * invalidPath("Path must start with /presentation or /slide[N]")
 */
export function invalidPath(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.INVALID_PATH, message, suggestion);
}

/**
 * Creates a NOT_FOUND error.
 *
 * @example
 * notFound("Slide", "5", "This presentation has only 3 slides")
 */
export function notFound(type: string, identifier: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.NOT_FOUND, `${type} ${identifier} not found`, suggestion);
}

/**
 * Creates an ALREADY_EXISTS error.
 *
 * @example
 * alreadyExists("Slide", "3", "Use --force to replace existing slides")
 */
export function alreadyExists(type: string, identifier: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.ALREADY_EXISTS, `${type} ${identifier} already exists`, suggestion);
}

/**
 * Creates an INVALID_INPUT error.
 *
 * @example
 * invalidInput("Slide index must be a positive integer", "Use slide[1] for the first slide")
 */
export function invalidInput(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.INVALID_INPUT, message, suggestion);
}

/**
 * Creates a NOT_SUPPORTED error.
 *
 * @example
 * notSupported("Adding charts via path is not yet implemented")
 */
export function notSupported(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.NOT_SUPPORTED, message, suggestion);
}
