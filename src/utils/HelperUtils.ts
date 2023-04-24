
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
    public static GenerateId(): string {
        const uniqueId = Date.now().toString(36) + Math.random().toString(36).substring(2);
        return uniqueId;
    }
}