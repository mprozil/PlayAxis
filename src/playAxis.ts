/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
     /**
     * Interface for viewmodel.
     *
     * @interface
     * @property {CategoryDataPoint[]} dataPoints - Set of data points the visual will render.
     */
    interface ViewModel {
        dataPoints: CategoryDataPoint[];
        settings: VisualSettings;
    };

    /**
     * Interface for data points.
     *
     * @interface
     * @property {string} category          - Corresponding category of data value.
     * @property {ISelectionId} selectionId - Id assigned to data point for cross filtering
     *                                        and visual interaction.
     */
    interface CategoryDataPoint {
        category: string;
        selectionId: ISelectionId;
    };
    
    /**
     * Interface for VisualChart settings.
     *
     * @interface
     * @property {{timeInterval:number}} transitionSettings - Object property that allows setting the time between transitions.
     * @property {{colorPicked:Fill}} colorSelector - Object property that allows setting the control buttons color.
     * @property {{show:boolean}} captionSettings - Object property that allows axis to be enabled.
     * @property {{captionColor:Fill}} captionSettings - Object property that allows setting the caption buttons.
     * @property {{fontSize:number}} captionSettings - Object property that allows setting the caption font size.
     */
    interface VisualSettings {        
        transitionSettings: {
            timeInterval: number;
        };
        colorSelector: {
            colorPicked: Fill;
        };
        captionSettings: {
            show: boolean;
            captionColor: Fill;
            fontSize: number;
        };
    }

    /**
     * Function that converts queried data into a view model that will be used by the visual.
     *
     * @function
     * @param {VisualUpdateOptions} options - Contains references to the size of the container
     *                                        and the dataView which contains all the data
     *                                        the visual had queried.
     * @param {IVisualHost} host            - Contains references to the host which contains services
     */
    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): ViewModel {
        let dataViews = options.dataViews;

        let defaultSettings: VisualSettings = {
            transitionSettings: {
                timeInterval: 1000,
            },
            colorSelector: {
                colorPicked: { solid: { color: "#000000" } },
            },
            captionSettings: {
                show: true,
                captionColor: { solid: { color: "#000000" } },
                fontSize: 16,
            }
        };

        let emptyViewModel: ViewModel = {
            dataPoints: [],
            settings: <VisualSettings>{}
        };

        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0].source)
            return emptyViewModel;

        let categorical = dataViews[0].categorical;
        let category = categorical.categories[0];

        let categoryDataPoints: CategoryDataPoint[] = [];
        
        let colorPalette: IColorPalette = host.colorPalette;
        let objects = dataViews[0].metadata.objects;

        let visualSettings: VisualSettings = {
            transitionSettings: {
                timeInterval: getValue<number>(objects, 'transitionSettings', 'timeInterval', defaultSettings.transitionSettings.timeInterval),
            },
            colorSelector: {
                colorPicked: getValue<Fill>(objects, 'colorSelector', 'colorPicked', defaultSettings.colorSelector.colorPicked),
            },
            captionSettings: {
                show: getValue<boolean>(objects, 'captionSettings', 'show', defaultSettings.captionSettings.show),
                captionColor: getValue<Fill>(objects, 'captionSettings', 'captionColor', defaultSettings.captionSettings.captionColor),
                fontSize: getValue<number>(objects, "captionSettings", "fontSize", defaultSettings.captionSettings.fontSize),
            }
        }

        for (let i = 0, len = Math.max(category.values.length); i < len; i++) {
            categoryDataPoints.push({
                category: category.values[i] + '',
                selectionId: host.createSelectionIdBuilder()
                    .withCategory(category, i)
                    .createSelectionId()
            });
        }
      
        return {
            dataPoints: categoryDataPoints,
            settings: visualSettings,
        };
    }

    enum Status {Play, Pause, Stop}

    export class Visual implements IVisual {
        private host: IVisualHost;
        private selectionManager: ISelectionManager;
        private svg: d3.Selection<SVGElement>;
        private visualDataPoints: CategoryDataPoint[];
        private visualSettings: VisualSettings;
        private status: Status;
        private lastSelected: number;
        private viewModel: ViewModel;
        private timers: any;

        constructor(options: VisualConstructorOptions) {
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();
            this.status = Status.Stop;
            this.timers = [];
            this.lastSelected = 0;            
           
            let buttonNames = ["play", "pause", "stop","previous","next"];
            let buttonPath = [
                    "M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm-3 17v-10l9 5.146-9 4.854z", 
                    "M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm-1 17h-3v-10h3v10zm5-10h-3v10h3v-10z", 
                    "M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm-1 17h-3.5v-10h9v10z",
                    "M22 12c0 5.514-4.486 10-10 10s-10-4.486-10-10 4.486-10 10-10 10 4.486 10 10zm-22 0c0 6.627 5.373 12 12 12s12-5.373 12-12-5.373-12-12-12-12 5.373-12 12zm13 0l5-4v8l-5-4zm-5 0l5-4v8l-5-4zm-2 4h2v-8h-2v8z",
                    "M12 2c5.514 0 10 4.486 10 10s-4.486 10-10 10-10-4.486-10-10 4.486-10 10-10zm0-2c-6.627 0-12 5.373-12 12s5.373 12 12 12 12-5.373 12-12-5.373-12-12-12zm-6 16v-8l5 4-5 4zm5 0v-8l5 4-5 4zm7-8h-2v8h2v-8z"
                   ];

            this.svg = d3.select(options.element).append("svg")
                 .attr("width","100%")
                 .attr("height","100%");

            for (let i = 0; i < buttonNames.length; ++i) {
                let container = this.svg.append('g')
                 .attr("transform","translate("+30*i+",0)")
                 .attr('class', "controls")
                 .attr('id', buttonNames[i]);           
                container.append("path")
                .attr("d", buttonPath[i]);
             }
            
            //Append caption text
            this.svg.append('text')
                .attr('dy','0.35em')
                .attr('id','label')
                .attr('transform', 'translate(150,12)');

            //Events on click
            this.svg.select("#play").on("click", () => {
                this.playAnimation();
            });
            this.svg.select("#stop").on("click", () => {
                this.stopAnimation();
            });
            this.svg.select("#pause").on("click", () => {
                this.pauseAnimation();
            });     
            this.svg.select("#previous").on("click", () => {
                this.step(-1);
            });     
            this.svg.select("#next").on("click", () => {
                this.step(1);
            });  

         }
         
        public update(options: VisualUpdateOptions) {
            let viewModel = this.viewModel = visualTransform(options, this.host);
            this.visualSettings = viewModel.settings;
            this.visualDataPoints = viewModel.dataPoints;           
            
            //Change colors
            let colorPicked = viewModel.settings.colorSelector.colorPicked.solid.color;
            let captionColor = viewModel.settings.captionSettings.captionColor.solid.color;
            this.svg.selectAll(".controls").attr("fill", colorPicked);
            this.svg.select("#label").attr("fill", captionColor);

            //Change caption font size
            let fontSize = viewModel.settings.captionSettings.fontSize;
            this.svg.select("#label").attr("font-size", fontSize);

            //Change title            
            if (this.visualSettings.captionSettings.show) {
                let title = options.dataViews[0].categorical.categories[0].source.displayName;           
                this.svg.select("#label").text(title);
                this.svg.attr("viewBox","0 0 260 24");
            } else {
                this.svg.select("#label").text("");
                this.svg.attr("viewBox","0 0 145 24");
            }
        }

        public playAnimation() {              
            if (this.status == Status.Play) return;

            this.svg.selectAll("#play, #next, #previous").attr("opacity", "0.3");
            this.svg.selectAll("#stop, #pause").attr("opacity", "1");

            let timeInterval = this.viewModel.settings.transitionSettings.timeInterval;
            let startingIndex = this.status == Status.Stop ? 0 : this.lastSelected + 1;
    
            for (let i = startingIndex; i < this.viewModel.dataPoints.length; ++i) {                           
                let timer = setTimeout(() => {
                    this.selectionManager.select(this.viewModel.dataPoints[i].selectionId);
                    this.lastSelected = i;
                    this.updateCaption(this.viewModel.dataPoints[i].category);
                }, (i - this.lastSelected) * timeInterval); 
                this.timers.push(timer);
            }
            this.status = Status.Play;
        }                

        public stopAnimation() {
            if (this.status == Status.Stop) return; 
            
            this.svg.selectAll("#pause, #stop, #next, #previous").attr("opacity", "0.3");
            this.svg.selectAll("#play").attr("opacity", "1");
            for (let i of this.timers) {
                clearTimeout(i);
            }
            this.updateCaption("");
            this.lastSelected = 0;
            this.selectionManager.clear();
            this.status = Status.Stop;
        }

        public pauseAnimation() {
            if (this.status == Status.Pause || this.status == Status.Stop) return;                                       

            this.svg.selectAll("#pause").attr("opacity", "0.3");
            this.svg.selectAll("#play, #stop, #next, #previous").attr("opacity", "1"); 
            for (let i of this.timers) {
                clearTimeout(i); 
            } 
            this.status = Status.Pause;
        }

        public step(step: number) {
            if (this.status == Status.Play || this.status == Status.Stop) return;                                       

            //Check if selection is within limits
            if ((this.lastSelected + step) < 0 || (this.lastSelected + step) > (this.viewModel.dataPoints.length-1)) return;

            this.lastSelected = this.lastSelected + step;
            this.selectionManager.select(this.viewModel.dataPoints[this.lastSelected].selectionId);
            this.updateCaption(this.viewModel.dataPoints[this.lastSelected].category);
            this.status = Status.Pause;
        }

        public updateCaption(caption: string) {
            if (this.visualSettings.captionSettings.show) {
                this.svg.select("#label").text(caption);
            }
        }
        /**
         * Enumerates through the objects defined in the capabilities and adds the properties to the format pane
         *
         * @function
         * @param {EnumerateVisualObjectInstancesOptions} options - Map of defined objects
         */
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            let objectName = options.objectName;
            let objectEnumeration: VisualObjectInstance[] = [];

            switch(objectName) {            
                case 'transitionSettings': 
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            timeInterval: this.visualSettings.transitionSettings.timeInterval
                        },
                        validValues: {
                            timeInterval: {
                                numberRange: {
                                    min: 500,
                                    max: 5000
                                }
                            }
                        },
                        selector: null
                    });
                break;
                case 'colorSelector':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            colorPicked: {
                                solid: {
                                    color: this.visualSettings.colorSelector.colorPicked.solid.color
                                }
                            }
                        },
                        selector: null
                    });
                break;
                case 'captionSettings':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            show: this.visualSettings.captionSettings.show,
                            captionColor: {
                                solid: {
                                    color: this.visualSettings.captionSettings.captionColor.solid.color
                                }
                            },
                            fontSize: this.visualSettings.captionSettings.fontSize
                        },
                        selector: null
                    });
                break;
            };
            return objectEnumeration;
        }
    }
}