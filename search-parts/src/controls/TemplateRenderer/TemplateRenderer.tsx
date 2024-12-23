import * as React from 'react';
import ITemplateRendererProps from './ITemplateRendererProps';
import ITemplateRendererState from './ITemplateRendererState';
import './TemplateRenderer.scss';
import { isEqual } from "@microsoft/sp-lodash-subset";
import * as DOMPurify from 'dompurify';
import { DomPurifyHelper } from '../../helpers/DomPurifyHelper';
import { ISearchResultsTemplateContext } from '../../models/common/ITemplateContext';
import { LayoutRenderType } from '@pnp/modern-search-extensibility';
import { Constants } from '../../common/Constants';
import * as ReactDOM from 'react-dom';
import { ITooltipHostStyles, TooltipDelay, TooltipHost } from '@fluentui/react/lib/Tooltip';
import { TooltipBasicExample } from "./TooltipBasicExample";
import { DirectionalHint, FontIcon } from '@fluentui/react';


// Need a root class to do not conflict with PnP Modern Search Styles.
const rootCssClassName = "pnp-modern-search";
const calloutProps = { gapSpace: 0 };

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
export class TemplateRenderer extends React.Component<ITemplateRendererProps, ITemplateRendererState> {

    private _domPurify: any;
    private _divTemplateRenderer: React.RefObject<HTMLDivElement>;

    constructor(props: ITemplateRendererProps) {
        super(props);

        this.state = {
        };

        this._domPurify = DOMPurify.default;

        this._domPurify.setConfig({
            ADD_TAGS: ['style', '#comment'],
            ADD_ATTR: ['target', 'loading'],
            ALLOW_DATA_ATTR: true,
            ALLOWED_URI_REGEXP: Constants.ALLOWED_URI_REGEXP,
            WHOLE_DOCUMENT: true,
        });

        this._domPurify.addHook('uponSanitizeElement', DomPurifyHelper.allowCustomComponentsHook);
        this._domPurify.addHook('uponSanitizeAttribute', DomPurifyHelper.allowCustomAttributesHook);

        // Create an instance of the div ref container 
        this._divTemplateRenderer = React.createRef<HTMLDivElement>();
    }

    public render() {
        return <div className={rootCssClassName} ref={this._divTemplateRenderer} />;
    }

    public async componentDidMount() {
        await this.updateTemplate(this.props);
    }

    public async componentDidUpdate(prevProps: ITemplateRendererProps) {

        if (!isEqual(prevProps.templateContent, this.props.templateContent) ||
            !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).inputQueryText, (this.props.templateContext as ISearchResultsTemplateContext).inputQueryText) ||
            !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).data, (this.props.templateContext as ISearchResultsTemplateContext).data) ||
            !isEqual(prevProps.templateContext.filters, this.props.templateContext.filters) ||
            !isEqual(prevProps.templateContext.properties, this.props.templateContext.properties) ||
            !isEqual(prevProps.templateContext.theme, this.props.templateContext.theme) ||
            !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).selectedKeys, (this.props.templateContext as ISearchResultsTemplateContext).selectedKeys)) {

            await this.updateTemplate(this.props);
        }
    }


    private async updateTemplate(props: ITemplateRendererProps): Promise<void> {
        let templateContent = props.templateContent;

        // Process the Handlebars template
        let template = await this.props.templateService.processTemplate(props.templateContext, templateContent, props.renderType);

        if (props.renderType == LayoutRenderType.Handlebars && typeof template === 'string') {

            // Sanitize the template HTML
            template = template ? this._domPurify.sanitize(`${template}`) : template;
            const templateAsHtml = new DOMParser().parseFromString(template as string, "text/html");

            // Add a random number at the right side of all <div class=
            const randomNumber = Math.floor(Math.random() * 1000);
            const randomNumberElement = templateAsHtml.createElement('div');
            randomNumberElement.textContent = `Random Number: ${randomNumber}`;
            randomNumberElement.style.float = 'inline-start';
            let tooltipProps: any = {
                onRenderContent: () => (
                    <ul style={{ margin: 10, padding: 0 }}>
                        <li><b>Description</b></li>

                    </ul>
                ),
            };

            const tooltipId = 'tooltip1234';
            const filterNameElements = templateAsHtml.querySelectorAll('.filter--name');
            filterNameElements.forEach((element) => {
                if (element) {
                    try {
                        // Create a container for the tooltip
                        const iconContainer = document.createElement('span');
                        iconContainer.style.display = 'inline-block';


                        ReactDOM.render(

                            // <TooltipBasicExample />

                            <TooltipHost
                                tooltipProps={tooltipProps}
                                content={"hi there"} // Tooltip text
                                calloutProps={{
                                    gapSpace: 0,
                                    isBeakVisible: true,
                                }}
                                styles={{
                                    root: { display: 'inline-block', position: 'relative' },
                                }}
                                delay={TooltipDelay.zero}
                                id={'tooltipId' + Math.random()}
                                directionalHint={DirectionalHint.rightCenter}
                            >
                                <FontIcon
                                    iconName="Info"
                                    aria-label="Info"
                                    style={{
                                        fontSize: 16,
                                        cursor: 'pointer',
                                        color: '#0078D4', // Optional: Add a color for the icon
                                    }}
                                />
                            </TooltipHost>
                            ,
                            iconContainer
                        );

                        // Append the container to the target element
                        element.appendChild(iconContainer);
                    } catch (error) {
                        console.error('Error rendering TooltipHost:', error);
                    }
                }
            });
            if (props.templateContext.properties.useMicrosoftGraphToolkit) {
                this.props.templateService.replaceDisambiguatedMgtElementNames(templateAsHtml);
            }

            // Get <style> tags from Handlebars template content and prefix all CSS rules by the Web Part instance ID to isolate styles
            const styleElements = templateAsHtml.getElementsByTagName("style");
            // let styles: string[] = [];
            // debugger;
            const allStyles: string[] = [];

            if (styleElements.length > 0) {

                // The prefix for all CSS selectors
                const elementPrefixId = `${this.props.templateService.TEMPLATE_ID_PREFIX}${this.props.instanceId}`;


                for (let i = 0; i < styleElements.length; i++) {
                    const style = styleElements.item(i);

                    if (style) {
                        let cssscope = style.dataset.cssscope as string;

                    if (cssscope !== undefined && cssscope === "layer") {

                        allStyles.push(`@layer { ${style.innerText} }`);
                    }
                    else {
                    }

                        allStyles.push(this.props.templateService.legacyStyleParser(style, elementPrefixId));

                    }
                }
            }

            if (this.props.templateContext.properties.useMicrosoftGraphToolkit && this.props.templateService.MgtCustomElementHelper.isDisambiguated) {
                allStyles.forEach((style, index) => {
                    allStyles[index] = style.replace(/mgt-/g, `${this.props.templateService.MgtCustomElementHelper.prefix}-`);
                });
            }

            if (this._divTemplateRenderer.current) {
                this._divTemplateRenderer.current.innerHTML = `<style>${allStyles.join(' ')}</style><div id="${this.props.templateService.TEMPLATE_ID_PREFIX}${this.props.instanceId}">${templateAsHtml.body.innerHTML}</div>`;
            }

        } else if (props.renderType === LayoutRenderType.AdaptiveCards && template instanceof HTMLElement) {

            if (this._divTemplateRenderer.current) {
                this._divTemplateRenderer.current.innerHTML = "";
                this._divTemplateRenderer.current.appendChild(template as HTMLElement);
            }
        }
    }
}
