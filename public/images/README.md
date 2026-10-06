# Site images

Static images for lessons and articles. Files here are served from the site root.

## URL pattern

| Content | Folder | Markdown |
| --- | --- | --- |
| Lesson | `public/images/lessons/{lesson-slug}/` | `![Alt text](/images/lessons/{lesson-slug}/file.webp)` |
| Article | `public/images/articles/{article-slug}/` | `![Alt text](/images/articles/{article-slug}/file.webp)` |

The `{slug}` matches the markdown filename without extension (for example `beginner-vbe-tour.md` → `beginner-vbe-tour/`).

## File guidelines

- Prefer **WebP** (PNG is fine for crisp UI chrome).
- Target **~1200px** max width; content column is ~768px (2x is enough for retina).
- Use **kebab-case** filenames: `developer-tab-visual-basic.webp`, not `Screenshot (1).png`.
- **Alt text** describes what the learner should notice, not "screenshot of Excel".

## Optional caption

```html
<figure class="lesson-figure">
  <img src="/images/lessons/beginner-vbe-tour/developer-tab-visual-basic.webp" alt="Developer tab with Visual Basic button" width="1200" height="675" loading="lazy" decoding="async" />
  <figcaption>Developer tab → Visual Basic opens the VBE.</figcaption>
</figure>
```

Plain markdown images work too; captions are optional.

## Lesson 1 (`beginner-vbe-tour/`) — suggested assets

Add these when ready (names are suggestions; keep paths in the lesson in sync):

| File | Use in lesson |
| --- | --- |
| `developer-tab-visual-basic.webp` | Open the VBE — ribbon |
| `customize-ribbon-developer.webp` | Enable Developer tab |
| `vbe-layout-project-explorer.webp` | Basic window explained |
| `vbe-edit-debug-toolbars.webp` | Edit and Debug toolbars |
| `insert-module-project-explorer.webp` | Insert a standard module |
| `code-window-module.webp` | Code window / resize |

After adding files, reference them in `src/content/lessons/beginner-vbe-tour.md`.
