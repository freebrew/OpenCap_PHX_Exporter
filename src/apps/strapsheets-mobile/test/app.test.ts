import { describe, expect, it } from "vitest";
import { addressProblem, parseAddresses } from "../src/Sessions";

describe("send-to addresses", () => {
  it("takes one or several addresses and rejects anything that is not one", () => {
    expect(parseAddresses(" Ann@Rig.com ")).toEqual(["ann@rig.com"]);
    expect(parseAddresses("a@x.com, b@y.com; a@x.com")).toEqual(["a@x.com", "b@y.com"]);
    expect(parseAddresses("")).toBeNull();
    expect(parseAddresses("a@x.com, nope")).toBeNull();
  });

  it("says which part is wrong, and that addresses have no spaces", () => {
    expect(addressProblem("bruno brottes@gmail.com")).toBe('"bruno" is not an address. Email addresses have no spaces; separate several with commas.');
    expect(addressProblem("a@x.com, nope")).toBe('"nope" is not an email address.');
    expect(addressProblem("bruno.brottes@gmail.com")).toBe("");
  });
});
