import { nameSchema } from "#/name-schema.ts";

export type Greeting = { message: string };

export function greet(name: string): Greeting {
  return { message: `Hello, ${nameSchema.parse(name)}` };
}
