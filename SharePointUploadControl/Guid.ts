export class Guid {
  public static newGuid(): Guid {
    return new Guid(
      "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
        const r = (Math.random() * 16) | 0;
        const v = c == "x" ? r : (r & 0x3) | 0x8;
        return v.toString(16);
      }),
    );
  }
  public static get empty(): string {
    return "00000000-0000-0000-0000-000000000000";
  }
  public get empty(): string {
    return Guid.empty;
  }
  public static isValid(str: string): boolean {
    const validRegex = /^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/;
    str = str.toLocaleLowerCase();
    return validRegex.test(str);
  }
  public static isEmpty(str: string): boolean {
    str = str.toLocaleLowerCase();
    return str === this.empty;
  }

  private value: string = this.empty;
  constructor(value?: string) {
    if (value) {
      if (Guid.isValid(value)) {
        this.value = value;
      }
    }
  }
  public toString(): string {
    return this.value;
  }

  public toJSON(): string {
    return this.value;
  }
}
