import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import { SPComponentLoader } from "@microsoft/sp-loader";

import * as jQuery from "jquery";
let Codebird: any = require("codebird");

import styles from "./SvTwitterFeedWebPart.module.scss";
import * as strings from "SvTwitterFeedWebPartStrings";

export interface ISvTwitterFeedWebPartProps {
  description: string;
  count: number;
}

export default class SvTwitterFeedWebPart extends BaseClientSideWebPart<ISvTwitterFeedWebPartProps> {

  public render(): void {

    let myCount: number = this.properties.count;

    //  font awesome is used for link icon (fa-external-link) and spinner (fa-spinner)
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.css");

    this.domElement.innerHTML = `
      <div class="${ styles.svTwitterFeed }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${ escape(this.properties.description) }</span>
              <input type="text" class="${ styles.input }" id="myInput"></input>
              <div class="${ styles.button }" id="myButton">Search</div>
              <div class="${ styles.loadingDiv }" id="loadingDiv"><i class="fa fa-spinner fa-spin spin-big"></i></div>
              <div class="${ styles.row }" id="dataContainer"></div>
            </div>
          </div>
        </div>
      </div>`;

    //  fire popDataContainer when enter is pressed in input field
    $("#myInput").keypress(function (e: any): void {
      let key: any = e.which;
      //  enter key is 13
      if (key === 13) {
        popDataContainer();
      }
    });

    //  fire popDataContainer when Search button is clicked
    jQuery("#myButton").click(function(): void {
      popDataContainer();
    });

    //  function to populate the dataContainer with the response from the Twitter API
    function popDataContainer(): void {

      //  clear search results
      jQuery("#dataContainer").empty();
      //  display the loading icon
      jQuery("#loadingDiv").css("display", "block");

      //  query Twitter App if search field is not empty/falsey
      if (jQuery("#myInput").val()) {
        //  codebird twitter API call
        let cb: any = new Codebird;
        //  these keys were registered to Twitter user @dikuw on 2018-02-20
        cb.setConsumerKey("9NoofVbuhwIvB79TtuXZNT8mX", "RKKE6mYdUba7WOklXDgrlOOHIOaSqlFeXYH9wYqWspS73Zr8oa");
        cb.setToken("43371090-XITgCw4mmcs3nRyuStrSfdrhVwjNf9sYevsNajMCs", "ueO1u2BAaHlUIWdwTOkjQBXdtyS0D3MaQuO9cArlNqbro");
        cb.__call(
          //  build search criteria
          //  available options are described here: https://developer.twitter.com/en/docs/tweets/search/api-reference/get-search-tweets
          "search_tweets", {
            q: jQuery("#myInput").val(),
            count: (myCount) ? myCount : 10
          },
          //  deal with response
          function(reply: any): void {
            //  hide the loading icon
            jQuery("#loadingDiv").css("display", "none");
            //  reply.statuses are the tweets
            let statuses: any = reply.statuses;
            //  for each tweet, append a div to the search results
            for (let i: number = 0; i < statuses.length; i++) {
              $("#dataContainer").append(`
                <div class="${ styles.status }">${ statuses[i].text }
                  <div class=${ styles.statusLink }>
                    <a href="https://twitter.com/${ statuses[i].user.screen_name }/status/${ statuses[i].id_str }" target="_blank">
                      <i class="fa fa-external-link"></i>
                    </a>
                  </div>
                </div>
                `
              );
            }
          }
        );
      } else {
        //  if search field is empty/falsey on input enter or search button click, display message for the user
        $("#dataContainer").html(`<div>Please enter a search string and try again.</div>`);
      }
    }

  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneSlider("count", {
                  label: "Number of tweets to return",
                  min: 1,
                  max: 100,
                  ariaLabel: "set number of returned tweets from 1 to 100"
                })
              ]
            }
          ]
        }
      ]
    };
  }

}
