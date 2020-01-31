import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { HttpClient } from '@microsoft/sp-http';
import { BaseSearchQueryModifier, IQuery, SearchQueryScenario } from '@microsoft/sp-search-extensibility';


/**
 * If your search query modifier uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISearchQueryModifierDemoProperties {
}

const AUTH_KEY: string = ''; // TODO: Add your own auth key

interface ITranslatorResponse {
  translations: [{
    text: string;
    to: string;
  }];
}

const LOG_SOURCE: string = 'SearchQueryModifierDemo';

export default class SearchQueryModifierDemo extends BaseSearchQueryModifier<ISearchQueryModifierDemoProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized SearchQueryModifierDemo');
    return Promise.resolve();
  }

  @override
  public modifySearchQuery(query: IQuery, scenario: SearchQueryScenario): Promise<IQuery> {
    let spanishText: string = '';
    let turkishText: string = '';

    return this.context.httpClient.post(
      'https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=en&to=es&to=tr', // tslint:disable-line: max-line-length
      HttpClient.configurations.v1,
      {
        body: JSON.stringify([{
          Text: query.queryText
        }]),
        headers: {
          'Ocp-Apim-Subscription-Key': AUTH_KEY,
          'Content-Type': 'application/json'
        }
      })
      .then(response => response.json())
      .then((json: ITranslatorResponse[]) => {
        json[0].translations.forEach(t => {
          if (t.to === 'es') { spanishText = t.text; }
          if (t.to === 'tr') { turkishText = t.text; }
        });
      })
      .then(() => {
        const queryText: string = `"${query.queryText}" OR "${spanishText}" OR "${turkishText}"`;
        console.log(`New query text is: "${queryText}"`);
        return Promise.resolve({ ...query, queryText });
      });
  }
}
