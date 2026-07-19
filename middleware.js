import { NextResponse } from "next/server";

export function middleware(req) {
  const auth = req.headers.get("authorization");

  const user = process.env.BASIC_AUTH_USER;
  const pass = process.env.BASIC_AUTH_PASS;

  if (auth) {
    const [, encoded] = auth.split(" ");
    const decoded = atob(encoded);
    const [u, p] = decoded.split(":");
    if (u === user && p === pass) {
      return NextResponse.next();
    }
  }

  return new NextResponse("Acesso restrito.", {
    status: 401,
    headers: { "WWW-Authenticate": 'Basic realm="MRM Personal Car"' },
  });
}

export const config = {
  matcher: "/((?!_next/static|_next/image|favicon.ico).*)",
};
