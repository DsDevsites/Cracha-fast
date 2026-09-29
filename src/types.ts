export type FieldKey='name'|'registration'|'role'
export type FieldPosition={x:number;y:number;width:number;height:number}
export type Person={id:string;name:string;registration:string;role:string;phone?:string;company?:string;manager?:string;trainingDate?:string}
export type BadgeTemplate={id:string;name:string;title:string;subtitle:string;primary:string;secondary:string;logoText:string;cardsPerPage:number;backgroundDataUrl?:string;fieldPositions?:Record<FieldKey,FieldPosition>}
export type Batch={id:string;name:string;templateId:string;people:Person[];createdAt:string}