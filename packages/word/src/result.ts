/**
 * Result envelope helpers for @officekit/word.
 *
 * All public API functions return a standardized Result<T> envelope:
 *   { ok: boolean, data?: T, error?: { code: string, message: string, suggestion?: string } }
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
 */
export function ok<T>(data: T): Result<T> {
  return { ok: true, data };
}

/**
 * Creates a successful result with undefined data (for void returns).
 */
export function okVoid(): Result<void> {
  return { ok: true };
}

/**
 * Creates a failed result.
 */
export function err(code: string, message: string, suggestion?: string): Result<never> {
  return {
    ok: false,
    error: suggestion ? { code, message, suggestion } : { code, message },
  };
}

/**
 * Creates a failed result from an error object.
 */
export function fail(error: Error, code = "unknown_error"): Result<never> {
  return { ok: false, error: { code, message: error.message } };
}

/**
 * Creates a failed result from an error code and message.
 */
export function failWith(code: string, message: string, suggestion?: string): Result<never> {
  return err(code, message, suggestion);
}

// ============================================================================
// Result Helpers
// ============================================================================

/**
 * Type guard to check if a result is successful.
 */
export function isOk<T>(result: Result<T>): result is Result<T> & { data: T } {
  return result.ok === true;
}

/**
 * Type guard to check if a result is a failure.
 */
export function isErr<T>(result: Result<T>): result is Result<T> & { error: ResultError } {
  return result.ok === false;
}

/**
 * Converts a nullable value to a Result.
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
 */
export function map<T, U>(result: Result<T>, fn: (data: T) => U): Result<U> {
  if (isErr(result)) {
    return result as Result<U>;
  }
  return ok(fn(result.data as T));
}

/**
 * Maps the error inside a result if failed.
 */
export function mapErr<T>(result: Result<T>, fn: (error: ResultError) => ResultError): Result<T> {
  if (isOk(result)) {
    return result;
  }
  return { ok: false, error: fn(result.error!) };
}

/**
 * Chains a result through a function that returns a result.
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
 */
export function getOrElse<T>(result: Result<T>, fallback: T): T {
  if (isErr(result)) {
    return fallback;
  }
  return result.data as T;
}

/**
 * Unwraps the data from a result, throwing if it's a failure.
 */
export function unwrap<T>(result: Result<T>): T {
  if (isErr(result)) {
    throw new Error(result.error.message);
  }
  return result.data as T;
}

/**
 * Unwraps the data from a result, returning a default value if it's a failure.
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

export const ErrorCodes = {
  INVALID_PATH: "invalid_path",
  NOT_FOUND: "not_found",
  ALREADY_EXISTS: "already_exists",
  NOT_SUPPORTED: "not_supported",
  INVALID_INPUT: "invalid_input",
  INVALID_FORMAT: "invalid_format",
  TIMEOUT: "timeout",
  PERMISSION_DENIED: "permission_denied",
  OPERATION_FAILED: "operation_failed",
  UNKNOWN_ERROR: "unknown_error",
  USAGE_ERROR: "usage_error",
} as const;

export type ErrorCode = (typeof ErrorCodes)[keyof typeof ErrorCodes];

// ============================================================================
// Error Factories
// ============================================================================

export function invalidPath(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.INVALID_PATH, message, suggestion);
}

export function notFound(type: string, identifier: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.NOT_FOUND, `${type} ${identifier} not found`, suggestion);
}

export function alreadyExists(type: string, identifier: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.ALREADY_EXISTS, `${type} ${identifier} already exists`, suggestion);
}

export function invalidInput(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.INVALID_INPUT, message, suggestion);
}

export function notSupported(message: string, suggestion?: string): Result<never> {
  return err(ErrorCodes.NOT_SUPPORTED, message, suggestion);
}
