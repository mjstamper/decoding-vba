# Decoding VBA

Source for [decodingvba.com](https://decodingvba.com) — free structured Excel VBA lessons and articles, with paid courses sold via external checkout.

## Stack

- [Astro 6](https://astro.build/) + [Tailwind CSS 4](https://tailwindcss.com/)
- Content in Markdown under `src/content/`
- Deployed to GitHub Pages via GitHub Actions

## Local development

```bash
npm install
npm run dev
```

Open `http://localhost:4321`.

Search uses Pagefind and only works after a build:

```bash
npm run build
npm run preview
```

Then open `/search/`.

If `npm install` fails with a certificate error on your network, try:

```powershell
$env:NODE_OPTIONS="--use-system-ca"
npm install
```

## Analytics

Production builds load [Plausible Analytics](https://plausible.io/) for `decodingvba.com`. Register the domain in your Plausible dashboard before deploying.

Copy [`.env.example`](.env.example) to `.env` if you need to override settings:

- `PUBLIC_PLAUSIBLE_DOMAIN` — defaults to `decodingvba.com`
- `PUBLIC_PLAUSIBLE_DISABLED=true` — turn off analytics in production builds

## Build

```bash
npm run build
npm run preview
```

## Project structure

```
src/
  config/site.ts       # Site name, nav, checkout placeholder
  content/
    lessons/           # Beginner / Intermediate / Advanced tracks
    articles/          # SEO guides
    courses/           # Paid course landing content
  pages/               # Routes
public/
  CNAME                # Custom domain for GitHub Pages
```

## Adding content

### Lesson

Create `src/content/lessons/my-lesson.md`:

```yaml
---
title: Lesson title
track: beginner          # beginner | intermediate | advanced
order: 9
summary: One-line description
draft: false
---
```

### Article

Create `src/content/articles/my-article.md` with `publishedDate` and `tags`.

### Course

Create `src/content/courses/my-course.md` with `checkoutUrl`, `price`, and `featured: true` for homepage promo.

Update the checkout URL in `src/config/site.ts` and the course frontmatter when Lemon Squeezy or Gumroad is ready.

## Deploy to GitHub Pages

1. Push this repo to GitHub (public repo for free Pages).
2. **Settings → Pages → Build and deployment → Source:** GitHub Actions.
3. Push to `main` — the workflow in `.github/workflows/deploy.yml` builds and deploys `dist/`.
4. **Settings → Pages → Custom domain:** `decodingvba.com`
5. Enable **Enforce HTTPS** after DNS verifies.

## Custom domain DNS

Replace `YOUR_GITHUB_USERNAME` with the account that owns the repo.

| Type  | Host | Value |
|-------|------|-------|
| A     | `@`  | `185.199.108.153` |
| A     | `@`  | `185.199.109.153` |
| A     | `@`  | `185.199.110.153` |
| A     | `@`  | `185.199.111.153` |
| CNAME | `www` | `YOUR_GITHUB_USERNAME.github.io` |

`public/CNAME` already contains `decodingvba.com`. DNS propagation can take up to 48 hours; often completes within 30 minutes.

### Optional www redirect

GitHub Pages custom domain settings handle apex vs www. Pick one canonical host in Astro (`site: 'https://decodingvba.com'`) and stick with it.

## Paid courses

- Do **not** commit paid videos, PDFs, or workbooks to this public repo.
- Host deliverables on Lemon Squeezy or Gumroad; link from course frontmatter `checkoutUrl`.
- Replace `https://decodingvba.lemonsqueezy.com/buy/placeholder` when the product is live.

## License

Content copyright Decoding VBA. Code in this repository is MIT unless noted otherwise.
