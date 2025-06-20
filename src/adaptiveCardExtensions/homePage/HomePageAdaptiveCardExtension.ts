import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HomePagePropertyPane } from './HomePagePropertyPane';

export interface IHomePageAdaptiveCardExtensionProps {
  title: string;
}

export interface IHomePageAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'HomePage_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HomePage_QUICK_VIEW';

export default class HomePageAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHomePageAdaptiveCardExtensionProps,
  IHomePageAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HomePagePropertyPane;

  public onInit(): Promise<void> {
    this.state = { };

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    // registers the quick view to open via QuickView action
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HomePage-property-pane'*/
      './HomePagePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HomePagePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
