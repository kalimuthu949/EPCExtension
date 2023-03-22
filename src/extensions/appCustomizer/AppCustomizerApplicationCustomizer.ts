import { Log } from "@microsoft/sp-core-library";
import styles from "./AppcustomizerApplicationCustomizer.module.scss";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";
import * as strings from "AppCustomizerApplicationCustomizerStrings";
import "./Custom.css";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import "@pnp/sp/attachments";
import "@pnp/sp/site-groups/web";

const LOG_SOURCE: string = "AppCustomizerApplicationCustomizer";

let isHome: boolean = false;

export interface IAppCustomizerApplicationCustomizerProperties {
  testMessage: string;
  Top: string;
  Bottom: string;
}

export default class AppCustomizerApplicationCustomizer extends BaseApplicationCustomizer<IAppCustomizerApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );
    super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
      });
      this.getCurrentPageDetails();
    });
    return Promise.resolve();
  }

  getCurrentPageDetails = async () => {
    await sp.web.lists
      .getByTitle("Site Pages")
      .items.get()
      .then((res) => {
        isHome = res.some(
          (li) =>
            li.Title.toLowerCase() == "intranet" &&
            li.ID == this.context.pageContext["_listItem"].id
        );
        if (!this._bottomPlaceholder && isHome) {
          this._bottomPlaceholder =
            this.context.placeholderProvider.tryCreateContent(
              PlaceholderName.Bottom,
              { onDispose: this._onDispose }
            );

          if (!this._bottomPlaceholder) {
            return;
          }

          if (this.properties) {
            let bottomString: string = this.properties.Bottom;
            if (!bottomString) {
              bottomString = "(Bottom property was not defined.)";
            }

            if (this._bottomPlaceholder.domElement) {
              this._bottomPlaceholder.domElement.innerHTML = `
                <div class="${styles.app}">
                    <div class="${styles.bottom}">
                      <a class="clsFooterRef" target="_blank" href="https://www.flyfrontier.com/green/">AMERICA'S &nbsp; <span class="spanContent">GREENEST</span> &nbsp; AIRLINE</a>
                    </div>
                </div>`;
            }
          }
        }
      })
      .catch((err) => {
        console.log(err);
      });
  };

  private _renderPlaceHolders(): void {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
      if (!this._topPlaceholder) {
        return;
      }

      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
              <div class="${styles.app}">
                  <div class="${styles.top}">
                  <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                    topString
                  )}
                  </div>
              </div>`;
        }
      }
    }
  }

  private _onDispose(): void {}
}
