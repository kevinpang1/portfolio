const PAGES = new Map();

function htmlResponse(body) {
  return new Response(body, {
    headers: {
      "content-type": "text/html; charset=utf-8",
      "cache-control": "public, max-age=300",
    },
  });
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    let pathname = decodeURIComponent(url.pathname);

    if (pathname === "/") {
      return Response.redirect(
        `${url.origin}/stock_holdings_percent_dashboard.html`,
        302,
      );
    }

    if (pathname.endsWith("/") && pathname.length > 1) {
      pathname = pathname.slice(0, -1);
    }

    const page = PAGES.get(pathname);
    if (page) {
      return htmlResponse(page);
    }

    return new Response("Not found", { status: 404 });
  },
};
