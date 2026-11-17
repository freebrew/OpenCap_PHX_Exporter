export type MailVendor = "google" | "microsoft";
export type EmailMode = "account" | "service" | "share";

export interface SendContext {
  online: boolean;
  signedIn: boolean;
  mailVendor: MailVendor | null;
  googleMailAllowed: boolean;
}

/** StrapSheets send order. The user's choice comes first; the share sheet is always last. */
export function sendOrder(mode: EmailMode, context: SendContext): EmailMode[] {
  const account = context.online && accountUsable(context);
  const service = context.online && context.signedIn;
  const order: EmailMode[] = [];
  switch (mode) {
    case "account":
      if (account) order.push("account");
      if (service) order.push("service");
      break;
    case "service":
      if (service) order.push("service");
      break;
    case "share":
      break;
    default: {
      const never: never = mode;
      return never;
    }
  }
  order.push("share");
  return order;
}

function accountUsable(context: SendContext): boolean {
  const vendor = context.mailVendor;
  switch (vendor) {
    case "google":
      return context.googleMailAllowed;
    case "microsoft":
      return true;
    case null:
      return false;
    default: {
      const never: never = vendor;
      return never;
    }
  }
}
