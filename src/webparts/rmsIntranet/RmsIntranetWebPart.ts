import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';

export interface IRmsIntranetWebPartProps {
  url: string;
  altura: number;
}

/**
 * Intranet RMS.
 *
 * O conteúdo da intranet vive num arquivo HTML autossuficiente, dentro de um
 * iframe — assim o CSS da RMS não colide com o do SharePoint e vice-versa.
 * A URL e a altura são configuráveis no painel de propriedades da web part.
 */
export default class RmsIntranetWebPart extends BaseClientSideWebPart<IRmsIntranetWebPartProps> {

  private static readonly URL_PADRAO = 'https://regis-monteiro.github.io/';
  private static readonly ALTURA_PADRAO = 1200;

  public render(): void {
    const url = (this.properties.url || RmsIntranetWebPart.URL_PADRAO).trim();
    const altura = Number(this.properties.altura) || RmsIntranetWebPart.ALTURA_PADRAO;

    this.domElement.innerHTML = `
      <div style="width:100%;background:#FBFAF7;border-radius:3px;overflow:hidden">
        <iframe
          src="${this.escapeAttr(url)}"
          title="Intranet RMS"
          style="width:100%;height:${altura}px;border:0;display:block"
          loading="lazy"
          referrerpolicy="no-referrer-when-downgrade"
          allow="clipboard-write; fullscreen"></iframe>
      </div>`;
  }

  /** Evita injeção pelo campo de URL do painel de propriedades. */
  private escapeAttr(valor: string): string {
    return valor
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Intranet RMS' },
          groups: [
            {
              groupName: 'Conteúdo',
              groupFields: [
                PropertyPaneTextField('url', {
                  label: 'Endereço da intranet',
                  description: 'URL do site publicado. Padrão: https://regis-monteiro.github.io/'
                }),
                PropertyPaneSlider('altura', {
                  label: 'Altura do quadro (px)',
                  min: 600,
                  max: 2400,
                  step: 50,
                  value: RmsIntranetWebPart.ALTURA_PADRAO
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
