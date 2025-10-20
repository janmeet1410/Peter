export interface IUserItem {
    name: string;
    mail: string;
    userPrincipalName: string;
    manager: string;
    givenName: string;
    title: string;
    mobilePhone: string;
    officeLocation: string;
    preferredLanguage: string;
    businessPhones: any;
    city: string;
    id: string;
    parent_id: string;
    photo:string;
}

export class ChartItem {

    public name: string;
    public mail: string;
    public userPrincipalName: string;
    public manager: string;  // manager user Property
    public givenName: string;
    public title: string;
    public mobilePhone: string;
    public officeLocation: string;
    public preferredLanguage: string;
    public businessPhones: any;
    public city: string;
    public id: string;
    public parent_id: string; // Importand propery to chart so naver delete this propery
    public url: string; // Importand propery to chart so naver delete this propery


    constructor(id: string, displayName: string, jobTitle: string, mail: string,manager: string,url:string) {
        this.id = id;
        this.name = displayName;
        this.title = jobTitle;
        this.mail = mail;
        this.url = url; // use same as, do not delete this or change this
        this.parent_id = manager; // use same as, do not delete this or change this
    }
}