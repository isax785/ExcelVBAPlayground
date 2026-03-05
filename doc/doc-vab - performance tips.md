# Performance & Interop Tips

*   **Minimize COM calls:** read/write ranges in blocks (arrays), not cell‑by‑cell.
*   **Avoid `Select/Activate`:** work on objects directly.
*   **Use `.Value2`** for speed/consistency.
*   **Turn off screen updating & events** around bulk ops; always restore in `CleanExit`.
*   **Named ranges & Tables (ListObjects):** stable references even as data grows.
*   **Interop:** When applicable, call Solver, Power Query refresh, or external tools (subject to policy).