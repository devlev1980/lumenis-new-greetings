import {BaseClientSideWebPart} from '@microsoft/sp-webpart-base';

export default class Item extends BaseClientSideWebPart<any> {
  public render(): void {
    this.domElement.innerHTML = `Hello`
  }
}
