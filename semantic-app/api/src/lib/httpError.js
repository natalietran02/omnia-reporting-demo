// A thin error type that carries the HTTP status it should produce, so every
// function's catch block can do the same `err instanceof HttpError ? err.status : 500`
// instead of each guessing at what went wrong.
class HttpError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

module.exports = { HttpError };
