declare namespace Cloudflare {
  interface Env {
    TOKEN_SIGNING_KEY: string;
  }
}

declare module "cloudflare:workers" {
  interface ProvidedEnv extends Cloudflare.Env {}
}
