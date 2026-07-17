# Web UI Guidelines

Use this as the default UI/UX standard for new web modules in this project.

## Wizard Layout

- Match the `DOCX Template Builder Web` module style.
- Use a centered `main` container with `max-width: 1280px`, `Times New Roman`, pale gray page background, and white bordered panels.
- Each module with multiple steps should use a horizontal step tab bar:
  - inactive step: light gray background
  - active step: blue background `#0f609b` with white text
  - left-aligned bold step labels
  - equal-width columns on desktop, stacked on small screens
  - compact fixed tab height matching `templates/configure.html`: about `46px`, `7px 8px` padding, `box-sizing: border-box`, no tall auto-expanding cells
- Step content should live in a `.panel` with white background, `1px #d9e2ec` border, `8px` radius, `16px` padding, and a stable minimum height around `560px`.

## Navigation

- Do not place primary `Trước` / `Tiếp` navigation inside the step body.
- Use a fixed footer navigation bar at the bottom of the viewport:
  - width: `min(1280px, calc(100vw - 40px))`
  - centered with `left: 50%` and `transform: translateX(-50%)`
  - bottom offset around `22px`
  - `Trước` pinned left with `margin-right: auto`
  - `Tiếp →` pinned right
- Hide `Trước` on the first step.
- Hide or replace `Tiếp` on the final step when there is no next step.
- Step-specific actions may change the `Tiếp` label, for example `Phân tích →` or `Sinh Excel →`, but the position stays fixed.

## Controls

- Keep inputs, selects, textareas, and buttons using the same typography and compact spacing as `templates/configure.html`.
- Use blue for primary actions, gray for secondary actions, and avoid introducing a separate visual language per module.
- Tables should use bordered cells and light gray headers consistent with the template builder.
