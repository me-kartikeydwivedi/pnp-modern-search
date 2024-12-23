import * as React from 'react';
import ITemplateRendererProps from './ITemplateRendererProps';
import './TemplateRenderer.scss';
import { isEqual } from "@microsoft/sp-lodash-subset";
import * as DOMPurify from 'dompurify';
import { DomPurifyHelper } from '../../helpers/DomPurifyHelper';
import { ISearchResultsTemplateContext } from '../../models/common/ITemplateContext';
import { LayoutRenderType } from '@pnp/modern-search-extensibility';
import { Constants } from '../../common/Constants';
import * as ReactDOM from 'react-dom';
import { ITooltipHostStyles, TooltipDelay, TooltipHost } from '@fluentui/react/lib/Tooltip';
import { DirectionalHint, FontIcon } from '@fluentui/react';
import { TooltipBasicExample } from "./TooltipBasicExample";

// Need a root class to do not conflict with PnP Modern Search Styles.
const rootCssClassName = "pnp-modern-search";
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

const TemplateRenderer: React.FC<ITemplateRendererProps> = (props) => {
    const [domPurify] = React.useState(() => {
        const purify = DOMPurify.default;
        purify.setConfig({
            ADD_TAGS: ['style', '#comment'],
            ADD_ATTR: ['target', 'loading'],
            ALLOW_DATA_ATTR: true,
            ALLOWED_URI_REGEXP: Constants.ALLOWED_URI_REGEXP,
            WHOLE_DOCUMENT: true,
        });
        purify.addHook('uponSanitizeElement', DomPurifyHelper.allowCustomComponentsHook);
        purify.addHook('uponSanitizeAttribute', DomPurifyHelper.allowCustomAttributesHook);
        return purify;
    });

    const divTemplateRenderer = React.useRef<HTMLDivElement>(null);

    React.useEffect(() => {
        const updateTemplate = async (props: ITemplateRendererProps) => {
            let templateContent = props.templateContent;

            // Process the Handlebars template
            let template = await props.templateService.processTemplate(props.templateContext, templateContent, props.renderType);

            if (props.renderType == LayoutRenderType.Handlebars && typeof template === 'string') {
                // Sanitize the template HTML
                template = template ? domPurify.sanitize(`${template}`) : template;
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
                                </TooltipHost>,
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
                    props.templateService.replaceDisambiguatedMgtElementNames(templateAsHtml);
                }

                // Get <style> tags from Handlebars template content and prefix all CSS rules by the Web Part instance ID to isolate styles
                const styleElements = templateAsHtml.getElementsByTagName("style");
                const allStyles:any = [];

                if (styleElements.length > 0) {
                    // The prefix for all CSS selectors
                    const elementPrefixId = `${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}`;

                    for (let i = 0; i < styleElements.length; i++) {
                        const style:any = styleElements.item(i);
                        let cssscope = style.dataset.cssscope as string;

                        if (cssscope !== undefined && cssscope === "layer") {
                            allStyles.push(`@layer { ${style.innerText} }`);
                        } else {
                            allStyles.push(props.templateService.legacyStyleParser(style, elementPrefixId));
                        }
                    }
                }

                if (props.templateContext.properties.useMicrosoftGraphToolkit && props.templateService.MgtCustomElementHelper.isDisambiguated) {
                    allStyles.forEach((style, index) => {
                        allStyles[index] = style.replace(/mgt-/g, `${props.templateService.MgtCustomElementHelper.prefix}-`);
                    });
                }

                if (divTemplateRenderer.current) {
                    divTemplateRenderer.current.innerHTML = `<style>${allStyles.join(' ')}</style><div id="${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}">${templateAsHtml.body.innerHTML}</div>`;
                }
            }
            
            else if (props.renderType == LayoutRenderType.AdaptiveCards && template instanceof HTMLElement) {
                if (divTemplateRenderer.current) {
                    divTemplateRenderer.current.innerHTML = "";
                    divTemplateRenderer.current.appendChild(template as HTMLElement);
                }
            }
        };

        updateTemplate(props);
    }, [props, domPurify]);

    // React.useEffect(() => {
    //     const updateTemplate = async (props: ITemplateRendererProps) => {
    //         let templateContent = props.templateContent;

    //         // Process the Handlebars template
    //         let template = await props.templateService.processTemplate(props.templateContext, templateContent, props.renderType);

    //         // if (props.renderType == LayoutRenderType.Handlebars && typeof template === 'string') {
    //         //     // Sanitize the template HTML
    //         //     template = template ? domPurify.sanitize(`${template}`) : template;
    //         //     const templateAsHtml = new DOMParser().parseFromString(template as string, "text/html");

    //         //     // Add a random number at the right side of all <div class=
    //         //     const randomNumber = Math.floor(Math.random() * 1000);
    //         //     const randomNumberElement = templateAsHtml.createElement('div');
    //         //     randomNumberElement.textContent = `Random Number: ${randomNumber}`;
    //         //     randomNumberElement.style.float = 'inline-start';
    //         //     let tooltipProps: any = {
    //         //         onRenderContent: () => (
    //         //             <ul style={{ margin: 10, padding: 0 }}>
    //         //                 <li><b>Description</b></li>
    //         //             </ul>
    //         //         ),
    //         //     };

    //         //     const tooltipId = 'tooltip1234';
    //         //     const filterNameElements = templateAsHtml.querySelectorAll('.filter--name');
    //         //     filterNameElements.forEach((element) => {
    //         //         if (element) {
    //         //             try {
    //         //                 // Create a container for the tooltip
    //         //                 const iconContainer = document.createElement('span');
    //         //                 iconContainer.style.display = 'inline-block';

    //         //                 ReactDOM.render(
    //         //                     // <TooltipHost
    //         //                     //     tooltipProps={tooltipProps}
    //         //                     //     content={"hi there"} // Tooltip text
    //         //                     //     calloutProps={{
    //         //                     //         gapSpace: 0,
    //         //                     //         isBeakVisible: true,
    //         //                     //     }}
    //         //                     //     styles={{
    //         //                     //         root: { display: 'inline-block', position: 'relative' },
    //         //                     //     }}
    //         //                     //     delay={TooltipDelay.zero}
    //         //                     //     id={'tooltipId' + Math.random()}
    //         //                     //     directionalHint={DirectionalHint.rightCenter}
    //         //                     // >
    //         //                     //     <FontIcon
    //         //                     //         iconName="Info"
    //         //                     //         aria-label="Info"
    //         //                     //         style={{
    //         //                     //             fontSize: 16,
    //         //                     //             cursor: 'pointer',
    //         //                     //             color: '#0078D4', // Optional: Add a color for the icon
    //         //                     //         }}
    //         //                     //     />
    //         //                     // </TooltipHost>
    //         //                   <TooltipBasicExample/>
    //         //                     ,
    //         //                     iconContainer
    //         //                 );

    //         //                 // Append the container to the target element
    //         //                 element.appendChild(iconContainer);
    //         //             } catch (error) {
    //         //                 console.error('Error rendering TooltipHost:', error);
    //         //             }
    //         //         }
    //         //     });

    //         //     if (props.templateContext.properties.useMicrosoftGraphToolkit) {
    //         //         props.templateService.replaceDisambiguatedMgtElementNames(templateAsHtml);
    //         //     }

    //         //     // Get <style> tags from Handlebars template content and prefix all CSS rules by the Web Part instance ID to isolate styles
    //         //     const styleElements = templateAsHtml.getElementsByTagName("style");
    //         //     const allStyles:any = [];

    //         //     if (styleElements.length > 0) {
    //         //         // The prefix for all CSS selectors
    //         //         const elementPrefixId = `${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}`;

    //         //         for (let i = 0; i < styleElements.length; i++) {
    //         //             const style:any = styleElements.item(i);
    //         //             let cssscope = style.dataset.cssscope as string;

    //         //             if (cssscope !== undefined && cssscope === "layer") {
    //         //                 allStyles.push(`@layer { ${style.innerText} }`);
    //         //             } else {
    //         //                 allStyles.push(props.templateService.legacyStyleParser(style, elementPrefixId));
    //         //             }
    //         //         }
    //         //     }

    //         //     if (props.templateContext.properties.useMicrosoftGraphToolkit && props.templateService.MgtCustomElementHelper.isDisambiguated) {
    //         //         allStyles.forEach((style, index) => {
    //         //             allStyles[index] = style.replace(/mgt-/g, `${props.templateService.MgtCustomElementHelper.prefix}-`);
    //         //         });
    //         //     }

    //         //     if (divTemplateRenderer.current) {
    //         //         divTemplateRenderer.current.innerHTML = `<style>${allStyles.join(' ')}</style><div id="${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}">${templateAsHtml.body.innerHTML}</div>`;
    //         //     }
    //         // }
    //         if (props.renderType == LayoutRenderType.Handlebars && typeof template === 'string') {
    //             // Sanitize the template HTML
    //             template = template ? domPurify.sanitize(`${template}`) : template;
    //             const templateAsHtml = new DOMParser().parseFromString(template as string, "text/html");

    //             // Add a random number at the right side of all <div class=
    //             const randomNumber = Math.floor(Math.random() * 1000);
    //             const randomNumberElement = templateAsHtml.createElement('div');
    //             randomNumberElement.textContent = `Random Number: ${randomNumber}`;
    //             randomNumberElement.style.float = 'inline-start';
    //             let tooltipProps: any = {
    //                 onRenderContent: () => (
    //                     <ul style={{ margin: 10, padding: 0 }}>
    //                         <li><b>Description</b></li>
    //                     </ul>
    //                 ),
    //             };

    //             const tooltipId = 'tooltip1234';
    //             const filterNameElements = templateAsHtml.querySelectorAll('.filter--name');
    //             filterNameElements.forEach((element) => {
    //                 if (element) {
    //                     try {
    //                         // Create a container for the tooltip
    //                         const iconContainer = document.createElement('span');
    //                         iconContainer.style.display = 'inline-block';

    //                         ReactDOM.render(
    //                             <TooltipHost
    //                                 tooltipProps={tooltipProps}
    //                                 content={"hi there"} // Tooltip text
    //                                 calloutProps={{
    //                                     gapSpace: 0,
    //                                     isBeakVisible: true,
    //                                 }}
    //                                 styles={{
    //                                     root: { display: 'inline-block', position: 'relative' },
    //                                 }}
    //                                 delay={TooltipDelay.zero}
    //                                 id={'tooltipId' + Math.random()}
    //                                 directionalHint={DirectionalHint.rightCenter}
    //                             >
    //                                 <FontIcon
    //                                     iconName="Info"
    //                                     aria-label="Info"
    //                                     style={{
    //                                         fontSize: 16,
    //                                         cursor: 'pointer',
    //                                         color: '#0078D4', // Optional: Add a color for the icon
    //                                     }}
    //                                 />
    //                             </TooltipHost>,
    //                             iconContainer
    //                         );

    //                         // Append the container to the target element
    //                         element.appendChild(iconContainer);
    //                     } catch (error) {
    //                         console.error('Error rendering TooltipHost:', error);
    //                     }
    //                 }
    //             });

    //             if (props.templateContext.properties.useMicrosoftGraphToolkit) {
    //                 props.templateService.replaceDisambiguatedMgtElementNames(templateAsHtml);
    //             }

    //             // Get <style> tags from Handlebars template content and prefix all CSS rules by the Web Part instance ID to isolate styles
    //             const styleElements = templateAsHtml.getElementsByTagName("style");
    //             const allStyles:any = [];

    //             if (styleElements.length > 0) {
    //                 // The prefix for all CSS selectors
    //                 const elementPrefixId = `${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}`;

    //                 for (let i = 0; i < styleElements.length; i++) {
    //                     const style:any = styleElements.item(i);
    //                     let cssscope = style.dataset.cssscope as string;

    //                     if (cssscope !== undefined && cssscope === "layer") {
    //                         allStyles.push(`@layer { ${style.innerText} }`);
    //                     } else {
    //                         allStyles.push(props.templateService.legacyStyleParser(style, elementPrefixId));
    //                     }
    //                 }
    //             }

    //             if (props.templateContext.properties.useMicrosoftGraphToolkit && props.templateService.MgtCustomElementHelper.isDisambiguated) {
    //                 allStyles.forEach((style, index) => {
    //                     allStyles[index] = style.replace(/mgt-/g, `${props.templateService.MgtCustomElementHelper.prefix}-`);
    //                 });
    //             }

    //             if (divTemplateRenderer.current) {
    //                 divTemplateRenderer.current.innerHTML = `<style>${allStyles.join(' ')}</style><div id="${props.templateService.TEMPLATE_ID_PREFIX}${props.instanceId}">${templateAsHtml.body.innerHTML}</div>`;
    //             }
    //         }
    //          else if (props.renderType == LayoutRenderType.AdaptiveCards && template instanceof HTMLElement) {
    //             if (divTemplateRenderer.current) {
    //                 divTemplateRenderer.current.innerHTML = "";
    //                 divTemplateRenderer.current.appendChild(template as HTMLElement);
    //             }
    //         }
    //     };
    //     const updateTemplateOnPropsChange = async (prevProps: ITemplateRendererProps) => {
    //         if (!isEqual(prevProps.templateContent, props.templateContent) ||
    //             !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).inputQueryText, (props.templateContext as ISearchResultsTemplateContext).inputQueryText) ||
    //             !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).data, (props.templateContext as ISearchResultsTemplateContext).data) ||
    //             !isEqual(prevProps.templateContext.filters, props.templateContext.filters) ||
    //             !isEqual(prevProps.templateContext.properties, props.templateContext.properties) ||
    //             !isEqual(prevProps.templateContext.theme, props.templateContext.theme) ||
    //             !isEqual((prevProps.templateContext as ISearchResultsTemplateContext).selectedKeys, (props.templateContext as ISearchResultsTemplateContext).selectedKeys)) {

    //             await updateTemplate(props);
    //         }
    //     };

    //     updateTemplateOnPropsChange(props);
    // }, [props]);

    return <div className={rootCssClassName} ref={divTemplateRenderer} />;
};

export default TemplateRenderer;
