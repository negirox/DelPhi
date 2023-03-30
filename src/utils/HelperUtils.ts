export class HelperUtils
{
    public static IsNullOrEmpty(obj:any):boolean{
        return obj === null && obj === undefined;
    }
    public static isEmpty(value: string): boolean {
        return value === undefined ||
          value === null ||
          value.length === 0;
    }
}