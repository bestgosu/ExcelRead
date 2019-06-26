declare function excelRead(filepath:string,rownum:number,startnum:number,endnum:number):Promise<any>;
declare function excelReadByBinaryString(bstr:string,rownum:number,startnum:number,endnum:number):Promise<any>;


export { excelRead, excelReadByBinaryString } 

