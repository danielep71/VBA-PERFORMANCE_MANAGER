# Project website

The public landing page is [`index.html`](index.html). It introduces the
project, makes its scope discoverable and links to the authoritative
documentation in the repository. It is not a software release, and it holds
no facts that the README does not.

## Publishing

Website changes follow the normal workflow: a pull request into the active
release branch, then the alignment of `main`. GitHub Pages publishes from
`main`:

- **Settings → Pages:** **Deploy from a branch**, branch **main**, folder
  **/docs**.
- **`.nojekyll`:** disables Jekyll processing. The page needs no build step,
  package or external font.
- **Script:** its only script, at the end of the page, starts the workbook
  download in a hidden frame. The links work without it.

The public address is <https://danielep71.github.io/VBA-PERFORMANCE_MANAGER/>.
After the first successful deployment:

1. set the repository's **About → Website** field to that address;
2. add the website link to the README navigation line.

The page uses only the images in [`assets/`](assets/README.md), where their
provenance is recorded. Other documentation links point to the rendered files
on GitHub, because Pages does not convert Markdown to HTML here.

## Maintenance

- **Alignment.** Keep the page's scope, release, certification facts and
  installation links aligned with the README.
- **At each release**, after publication, update:
  - the release named in the hero, the status line and the resources;
  - the certification line in the Assurance section;
  - the workbook download link. From v1.4.1 the workbook ships inside
    `VBA-Performance-Manager-vX.Y.Z.zip` (#72), so the link and its label
    change. Check that it still starts a download, and keep its
    `data-direct-download` attribute.
- **Counts.** Do not copy counts that change between releases, such as the
  number of static checks. Certified figures are quoted only for a named
  release.
- **Addresses.** Keep the canonical URL, the Open Graph URL and image, and
  [`sitemap.xml`](sitemap.xml) aligned if the repository or hosting address
  changes. The sitemap lists only the landing page.
- **Checks before merge.** Run the static gate as for any change; website-only
  changes need no Excel run. Check:
  - desktop and narrow-screen layout;
  - keyboard navigation;
  - local anchors and image paths;
  - external destinations;
  - that the page works without JavaScript.

## Search engines

Publishing does not submit the site to search engines. To use Google Search
Console:

1. add the site as a URL-prefix property;
2. add the exact HTML verification tag supplied by the maintainer to the
   page's `head`, through the normal pull-request workflow. Never insert an
   invented token;
3. once verification succeeds, submit `sitemap.xml`.

Do not add a project-level `robots.txt`: crawlers read only the host-root
one.
