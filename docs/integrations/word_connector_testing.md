# Word Connector Testing & Troubleshooting

## CORS Configuration

The Word Connector task pane (run locally via `https://localhost:3000`) requires explicit CORS permission to call the HaloBridge API.

### Required Settings
The following origin must be allowed in the backend environment:
- `https://localhost:3000`
- `https://127.0.0.1:3000` (optional fallback)

In the `.env` file, ensure `HALOBRIDGE_WORD_CONNECTOR_ALLOWED_ORIGINS` includes these values.

### Common Failure: ERR_NETWORK
If you see `ERR_NETWORK` or `Network/CORS error: Request blocked before response` in the Word task pane console, it usually means:
1. **CORS Blocked**: The backend `Access-Control-Allow-Origin` header does not match `https://localhost:3000`.
2. **Preflight Failure**: The `OPTIONS` request failed or returned headers that didn't allow `Authorization` or `Content-Type`.
3. **Invalid Certificate**: Since the add-in runs on `https`, the backend must also be on `https` (or standard browser security rules will block the 'mixed content' or insecure request).

## Diagnostic Steps

1. **Check Server Logs**: Look for `[CORS] Blocked request from unauthorized origin` in the terminal.
2. **Inspect Network Tab**: Look for the `OPTIONS` preflight request. Ensure it returns `200 OK` or `204 No Content` with appropriate CORS headers.
3. **Verify Auth**: Ensure the `Authorization` header is being sent correctly and is allowed in `Access-Control-Allow-Headers`.
