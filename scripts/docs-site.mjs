const configured = process.env.NORI_DOCS_BASE ?? "/";
if (!/^\/(?:[A-Za-z0-9_-]+\/)*$/.test(configured))
  throw Error(
    "NORI_DOCS_BASE must be a slash-delimited site path, such as /nori/",
  );
export const docsBase = configured;
export const sitePath = (path) => docsBase + path.replace(/^\//, "");
