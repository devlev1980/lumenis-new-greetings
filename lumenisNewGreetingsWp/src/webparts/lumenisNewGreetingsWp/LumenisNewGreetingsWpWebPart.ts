import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './LumenisNewGreetingsWpWebPart.module.scss';
import * as strings from 'LumenisNewGreetingsWpWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as pnp from 'sp-pnp-js';
import { Web } from "sp-pnp-js";
require('jquery');
import * as $ from 'jquery';

require('../lumenisNewGreetingsWp/LumenisNewGreetingsWpWebPart.scss')



export interface ILumenisNewGreetingsWpWebPartProps {
  description: string;
  listName: string;
  webUrl: string;
  wpTitle: string;
  sendYourGreetings: string;
  separator: string;
  daysBefore: string;
  daysAfter: string;
  displayDate: boolean;
}

export default class LumenisNewGreetingsWpWebPart extends BaseClientSideWebPart <ILumenisNewGreetingsWpWebPartProps> {

  private _newEmployeesListName = 'NewEmployees';

    public render(): void {
        let d = new Date();
        let thisTime= d.getTime();

        // SPComponentLoader.loadCss("/Style Library/IMF.O365.Lumenis/css/greetings.css");
        // SPComponentLoader.loadCss('./LumenisNewGreetingsWpWebPart.module.scss');
        this.domElement.innerHTML = `
        <p id="LumenisGreetingsWpWebPartID" class="anchorinpage"></p>
      <div class="event_wrapper" style="padding: 0px; margin-top: 0vw;" id="greetingsWP">
        <h3 class="WPtitle">${escape(this.properties.wpTitle)}</h3>
        <div class="container WPevents" id="WPevents${thisTime}">
            <div class="scrollEvents" id="scrollEvents${thisTime}"></div>

        </div>
    </div>`;
    let greetingsList = this.properties.listName;
      let query = this.Build(greetingsList);
      let webUrl = this.properties.webUrl;
      let sendYourGreetings = this.properties.sendYourGreetings;
        let separator = this.properties.separator;
      let absUrl = this.context.pageContext.site.absoluteUrl;
      let web = pnp.sp.web;
      if (webUrl != "") {
          web = new Web(absUrl + webUrl);
      }
      else {
          web = new Web(absUrl);
      }
      web.get().then((w) => {

          const q: any = {
              ViewXml: query
          };

          web.lists.getByTitle(greetingsList).getItemsByCAMLQuery(q).then((r:any[])=>{
            var _greetingsList = greetingsList;
            r.forEach((result) => {
                let _listName = _greetingsList;
                let _isNewEmployee = _greetingsList == this._newEmployeesListName;
                let itemTitle = result.Title;
                let itemRole;
                let itemEventType = result.EventType;
                let itemEventAuthorId = result.EventAuthorId;
                let itemMonth = result.EventMonth;
                let itemDay = result.EventDay;
                let babyGender = result.BabyGender;
                if (itemEventType == null) {
                    itemEventType = "birthday";
                }
                let mainID = result.ID;

                let userEmail = "#";
                let userPosition = "";

                let eventImg = "";
                let eventArrowImg = "";
                switch (itemEventType) {
                    case "birthday":
                        eventImg = "birthdayIcon.png";
                        itemRole = `${itemDay}${separator}${itemMonth}  מזל טוב`;
                        eventArrowImg = "pinkArrow.jpg";
                        break;
                    case "newborn":
                        eventImg = "handshakeIcon.png";
                        itemRole = "ברוך הבא ללומניס";
                        eventArrowImg = "darkBlueArrow.jpg";
                        break;
                    case "wedding":
                        eventImg = "weddingIcon.png";
                        itemRole = "מזל טוב לנישואיך";
                        eventArrowImg = "yellowArrow.jpg";
                        break;
                    case "baby":
                        eventImg = "strollerIcon.jpeg";
                        eventArrowImg = "lightBlueArrow.jpg";
                        switch (babyGender) {
                            case "בן":
                                itemRole = "מזל טוב להולדת הבן";
                                break;
                            case "בת":
                                itemRole = "מזל טוב להולדת הבת";
                                break;
                        }
                        break;
                }

                let selectContainer: Element = this.domElement.querySelector(`#scrollEvents${thisTime}`);

                if(this.properties.displayDate){
                selectContainer.insertAdjacentHTML("beforeend",
                `<div class="grtmpitem eventItemNews ${itemEventType}" userId="${itemEventAuthorId}" useremail="" id=${mainID}>

                <div class="celebrant_detailes">
                    <div class="evt_ic">
                        <img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}">
                    </div>
                    <div class="info">
                    <div class="user_name"></div>
                     <span class="user_office">${itemRole}</span>

</div>
<!--<div class="eventDate">-->
                    <div class="evt_ic">
                        <img src="/Style Library/IMF.O365.Lumenis/img/${eventArrowImg}">
<!--                   <div> <img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}"></div>-->
                     <a class="wish_btn" href="" title=""><br>${sendYourGreetings}</a>
                    </div>
                </div>
                </div>

                </div>`);
                }else{
                    selectContainer.insertAdjacentHTML("beforeend",
                    `<div class="grtmpitem eventItemNews ${itemEventType}" userId="${itemEventAuthorId}" useremail="" id=${mainID}>

                    <div class="celebrant_detailes">
                        <div class="user_name"></div>
                        <div class="user_office">${itemRole}</div>
                        <a class="wish_btn" href="" title="">${sendYourGreetings}</a>
                    </div>
                    <div class="eventDate">
                        <div class="evt_ic"><img src="/Style Library/IMF.O365.Lumenis/img/${eventImg}"></div>
                    </div>
                    </div>`);

                }
                var _picture = _isNewEmployee && result.Picture ? result.Picture : undefined;

                if(itemEventAuthorId){
                    web.siteUsers.getById(itemEventAuthorId).get().then(function(result){(function (result, picture) {
                        // console.log(mainID);
                        // console.log(result);
                        $(`#WPevents${thisTime} .eventItemNews#${mainID}`).attr('useremail', result.Email);
                        $(`#WPevents${thisTime} .eventItemNews#${mainID} .user_pic_date img`).attr('src', picture && picture.Url? picture.Url : "/_vti_bin/DelveApi.ashx/people/profileimage?userId=" +  result.Email);
                        $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .user_name`).text(result.Title);
                        $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .wish_btn`).attr('href', 'mailto:' + result.Email);
                    })(result, _picture)});
                }else if(_isNewEmployee && _picture && _picture.Url){
                    $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .user_name`).text(itemTitle);
                    $(`#WPevents${thisTime} .eventItemNews#${mainID} .user_pic_date img`).attr('src', _picture.Url);
                    $(`#WPevents${thisTime} .eventItemNews#${mainID} .celebrant_detailes .wish_btn`).css('visibility', 'hidden');
                }
            });
        });

      }).catch((err) => {
          console.log(err);
      });
    }

  public Build(listName): string {

      let eventDay = parseInt(moment().add(-(parseInt(this.properties.daysBefore)), 'd').format("DD"));
      let eventMonth = parseInt(moment().add(-(parseInt(this.properties.daysBefore)), 'd').format("MM"));

      let nextDate = moment().add((parseInt(this.properties.daysAfter)), 'd');
      let nextEventMonth = parseInt(nextDate.format("MM"));
      let nextEventDay = parseInt(nextDate.format("DD"));


      let firstMonthDays = [];
      let secondMonthDays = [];

      //next date is on the same month
      if (nextEventMonth == eventMonth) {
          let _i1 = 0;
          for (let i1 = eventDay; i1 <= nextEventDay; i1++) {
              firstMonthDays[_i1] = i1;
              _i1 = _i1 + 1;
          }
      }
      else {

          //next date is on next month
          let _i = 0;
          let _j = 0;
          let tempDate = moment().add(-(parseInt(this.properties.daysBefore)), 'd');
          while (tempDate.isSame(nextDate) || tempDate.isBefore(nextDate)) {
              if (parseInt(tempDate.format("MM")) == eventMonth) {
                  firstMonthDays[_i] = parseInt(tempDate.format("DD"));
                  _i += 1;
              }
              else {
                  secondMonthDays[_j] = parseInt(tempDate.format("DD"));
                  _j += 1;
              }

              tempDate = tempDate.add(1, 'd');
          }
      }

      let _tempQuery = '';

      if (secondMonthDays.length == 0) {

          //build query for current month
          let _in = '';
          for (let i = 0; i < firstMonthDays.length; i++) {
              _in += '<Value Type="Number">' + firstMonthDays[i] + '</Value>';
          }
          _tempQuery =
              '<And>' +
              '<And>' +
              '<Eq>' +
              '<FieldRef Name="EventMonth" />' +
              '<Value Type="Number">' + eventMonth + '</Value>' +
              '</Eq>' +
              '<In>' +
              '<FieldRef Name="EventDay" />' +
              '<Values>' + _in + '</Values>' +
              '</In>' +
              '</And>' +
              '<Or>' +
              '<Or>' +
              '<IsNull>' +
              '<FieldRef Name="Expires" />' +
              '</IsNull>' +
              '<And>' +
              '<IsNull>' +
              '<FieldRef Name="EventType" />' +
              '</IsNull>' +
              '<Eq>' +
              '<FieldRef Name="EventType" />' +
              '<Value Type="Text">birthday</Value>' +
              '</Eq>' +
              '</And>' +
              '</Or>' +
              '<Gt>' +
              '<FieldRef Name="Expires" />' +
              '<Value Type="DateTime">' +
              '<Today />' +
              '</Value>' +
              '</Gt>' +
              '</Or>' +
              '</And>';

      }
      else {

          //build query for 2 months
          let _in_first = '';
          let _in_second = '';

          for (let i3 = 0; i3 < firstMonthDays.length; i3++) {
              _in_first += '<Value Type="Number">' + firstMonthDays[i3] + '</Value>';
          }
          for (let i4 = 0; i4 < secondMonthDays.length; i4++) {
              _in_second += '<Value Type="Number">' + secondMonthDays[i4] + '</Value>';
          }
          _tempQuery =
              '<Or>' +
              '<And>' +
              '<And>' +
              '<Eq>' +
              '<FieldRef Name="EventMonth" />' +
              '<Value Type="Number">' + eventMonth + '</Value>' +
              '</Eq>' +
              '<In>' +
              '<FieldRef Name="EventDay" />' +
              '<Values>' + _in_first + '</Values>' +
              '</In>' +
              '</And>' +
              '<Or>' +
              '<Or>' +
              '<IsNull>' +
              '<FieldRef Name="Expires" />' +
              '</IsNull>' +
              '<And>' +
              '<IsNull>' +
              '<FieldRef Name="EventType" />' +
              '</IsNull>' +
              '<Eq>' +
              '<FieldRef Name="EventType" />' +
              '<Value Type="Text">birthday</Value>' +
              '</Eq>' +
              '</And>' +
              '</Or>' +
              '<Gt>' +
              '<FieldRef Name="Expires" />' +
              '<Value Type="DateTime">' +
              '<Today />' +
              '</Value>' +
              '</Gt>' +
              '</Or>' +
              '</And>' +
              '<And>' +
              '<And>' +
              '<Eq>' +
              '<FieldRef Name="EventMonth" />' +
              '<Value Type="Number">' + nextEventMonth + '</Value>' +
              '</Eq>' +
              '<In>' +
              '<FieldRef Name="EventDay" />' +
              '<Values>' + _in_second + '</Values>' +
              '</In>' +
              '</And>' +
              '<Or>' +
              '<Or>' +
              '<IsNull>' +
              '<FieldRef Name="Expires" />' +
              '</IsNull>' +
              '<And>' +
              '<IsNull>' +
              '<FieldRef Name="EventType" />' +
              '</IsNull>' +
              '<Eq>' +
              '<FieldRef Name="EventType" />' +
              '<Value Type="Text">birthday</Value>' +
              '</Eq>' +
              '</And>' +
              '</Or>' +
              '<Gt>' +
              '<FieldRef Name="Expires" />' +
              '<Value Type="DateTime">' +
              '<Today />' +
              '</Value>' +
              '</Gt>' +
              '</Or>' +
              '</And>' +
              '</Or>';
      }

      //get items for currend date +- 3 days
      let tempQuery = '<View><ViewFields>' +
          '<FieldRef Name="ID"/><FieldRef Name="Title"/><FieldRef Name="EventType"/><FieldRef Name="BabyGender"/><FieldRef Name="EventAuthor"/><FieldRef Name="EventMonth"/><FieldRef Name="EventDay"/><FieldRef Name="Role"/><FieldRef Name="Expires"/>' +
          (listName == this._newEmployeesListName ? '<FieldRef Name="Picture"/>' : '') +
          '</ViewFields>' +
          '<Query>' +
          '<Where>' + _tempQuery +
          '</Where>' +
          '<OrderBy>' +
          '<FieldRef Name="EventMonth" Ascending="TRUE" />' +
          '<FieldRef Name="EventDay" Ascending="TRUE" />' +
          '</OrderBy>' +
          '</Query>' +
          '<RowLimit>10000</RowLimit></View>';


      return tempQuery;
  }

  // @ts-ignore
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                  }),
                PropertyPaneTextField('webUrl', {
                    label: "Web url"
                }),
                  PropertyPaneTextField('listName', {
                      label: "List name"
                  }),
                  PropertyPaneTextField('wpTitle', {
                      label: "Wp title"
                  }),
                  PropertyPaneTextField('sendYourGreetings', {
                      label: "Send your greetings"
                  }),
                  PropertyPaneTextField('separator', {
                      label: "Separator"
                  }),
                  PropertyPaneTextField('daysBefore', {
                      label: "Days before"
                  }),
                  PropertyPaneCheckbox('displayDate', {
                    text: "Display date"
                }),
                  PropertyPaneTextField('daysAfter', {
                      label: "Days after"
                  })
              ]
            }
          ]
        }
      ]
    };
  }
}
