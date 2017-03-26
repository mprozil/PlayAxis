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
     * @property {{autoStart:boolean}} transitionSettings - Object property to enable or disable auto start option.
     * @property {{loop:boolean}} transitionSettings - Object property to enable or disable loop option.
     * @property {{timeInterval:number}} transitionSettings - Object property that allows setting the time between transitions.
     * @property {{pickedColor:Fill}} colorSelector - Object property that allows setting the control buttons color.
     * @property {{showAll:boolean}} colorSelector - Object property to enable or disable individual colors.
     * @property {{playColor:Fill}} colorSelector - Object property that allows setting the color for play button.
     * @property {{pauseColor:Fill}} colorSelector - Object property that allows setting the color for pause button.
     * @property {{stopColor:Fill}} colorSelector - Object property that allows setting the color for stop button..
     * @property {{previousColor:Fill}} colorSelector - Object property that allows setting the color the previous button.
     * @property {{nextColor:Fill}} colorSelector - Object property that allows setting the color for next button.
     * @property {{show:boolean}} captionSettings - Object property that allows axis to be enabled.
     * @property {{captionColor:Fill}} captionSettings - Object property that allows setting the caption buttons.
     * @property {{fontSize:number}} captionSettings - Object property that allows setting the caption font size.
     */
    interface VisualSettings {        
        transitionSettings: {
            autoStart: boolean;
            loop: boolean;
            timeInterval: number;
        };
        colorSelector: {
            pickedColor: Fill;
            showAll: boolean;
            playColor: Fill;
            pauseColor: Fill;
            stopColor: Fill;
            previousColor: Fill;
            nextColor: Fill;
        };
        captionSettings: {
            show: boolean;
            captionColor: Fill;
            fontSize: number;
            align: string;
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
                autoStart: false,
                loop: false,
                timeInterval: 1000,
            },
            colorSelector: {
                pickedColor: { solid: { color: "#000000" } },
                showAll: false,
                playColor: { solid: { color: "#f2c811" } },
                pauseColor: { solid: { color: "#1769b8" } },
                stopColor: { solid: { color: "#f42550" } },
                previousColor: { solid: { color: "#12b159" } },
                nextColor: { solid: { color: "#a81de8" } },
            },
            captionSettings: {
                show: true,
                captionColor: { solid: { color: "#000000" } },
                fontSize: 16,
                align: "left",
            }
        };

        let categorical = dataViews[0].categorical;
        let category = categorical.categories[0];

        let categoryDataPoints: CategoryDataPoint[] = [];
        
        let colorPalette: IColorPalette = host.colorPalette;
        let objects = dataViews[0].metadata.objects;

        let visualSettings: VisualSettings = {
            transitionSettings: {
                autoStart: getValue<boolean>(objects, 'transitionSettings', 'autoStart', defaultSettings.transitionSettings.autoStart),
                loop: getValue<boolean>(objects, 'transitionSettings', 'loop', defaultSettings.transitionSettings.loop),
                timeInterval: getValue<number>(objects, 'transitionSettings', 'timeInterval', defaultSettings.transitionSettings.timeInterval),
            },
            colorSelector: {
                pickedColor: getValue<Fill>(objects, 'colorSelector', 'pickedColor', defaultSettings.colorSelector.pickedColor),
                showAll: getValue<boolean>(objects, 'colorSelector', 'showAll', defaultSettings.colorSelector.showAll),
                playColor: getValue<Fill>(objects, 'colorSelector', 'playColor', defaultSettings.colorSelector.playColor),
                pauseColor: getValue<Fill>(objects, 'colorSelector', 'pauseColor', defaultSettings.colorSelector.pauseColor),
                stopColor: getValue<Fill>(objects, 'colorSelector', 'stopColor', defaultSettings.colorSelector.stopColor),
                previousColor: getValue<Fill>(objects, 'colorSelector', 'previousColor', defaultSettings.colorSelector.previousColor),
                nextColor: getValue<Fill>(objects, 'colorSelector', 'nextColor', defaultSettings.colorSelector.nextColor),
            },
            captionSettings: {
                show: getValue<boolean>(objects, 'captionSettings', 'show', defaultSettings.captionSettings.show),
                captionColor: getValue<Fill>(objects, 'captionSettings', 'captionColor', defaultSettings.captionSettings.captionColor),
                fontSize: getValue<number>(objects, "captionSettings", "fontSize", defaultSettings.captionSettings.fontSize),
                align: getValue<string>(objects, "captionSettings", "align", defaultSettings.captionSettings.align),
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

    /**
     * Function that checks if data is ready to be used by the visual.
     *
     * @function
     * @param {VisualUpdateOptions} options - Contains references to the size of the container
     *                                        and the dataView which contains all the data
     *                                        the visual had queried.
     */
    function isDataReady(options: VisualUpdateOptions) {
        if(!options
        || !options.dataViews
        || !options.dataViews[0]
        || !options.dataViews[0].categorical
        || !options.dataViews[0].categorical.categories
        || !options.dataViews[0].categorical.categories[0].source)
        {
            return false;
        }

        return true;                         
    }    

    enum Status {Play, Pause, Stop}

    export class Visual implements IVisual {
        private host: IVisualHost;
        private selectionManager: ISelectionManager;
        private svg: d3.Selection<SVGElement>;
        private controlsSVG: d3.Selection<SVGElement>;
        private captionSVG: d3.Selection<SVGElement>;
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

            this.controlsSVG = this.svg.append('svg');
            for (let i = 0; i < buttonNames.length; ++i) {
                let container = this.controlsSVG.append('g')
                 .attr('class', "controls")
                 .attr('transform','translate(' + 30*i + ')')
                 .attr('id', buttonNames[i]); 
                container.append("path")
                .attr("d", buttonPath[i]);
             }
            
            //Append caption text           
            this.captionSVG = this.svg.append('svg');
            let captionBox = this.captionSVG.append('g');
            captionBox.append('text')
                .attr('dominant-baseline','middle')
                .attr("y","14")
                .attr('id','label');

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

            if (isDataReady(options) == false) {
                return;
            }

            this.stopAnimation();
            let viewModel = this.viewModel = visualTransform(options, this.host);
            this.visualSettings = viewModel.settings;
            this.visualDataPoints = viewModel.dataPoints;        

            //Start playing without click 
            if (this.visualSettings.transitionSettings.autoStart) { 
                this.playAnimation();
            }

            //Change colors         
            if (this.visualSettings.colorSelector.showAll) {
                let playColor = viewModel.settings.colorSelector.playColor.solid.color;
                let pauseColor = viewModel.settings.colorSelector.pauseColor.solid.color;
                let stopColor = viewModel.settings.colorSelector.stopColor.solid.color;
                let previousColor = viewModel.settings.colorSelector.previousColor.solid.color;
                let nextColor = viewModel.settings.colorSelector.nextColor.solid.color;
                this.svg.selectAll("#play").attr("fill", viewModel.settings.colorSelector.playColor.solid.color);
                this.svg.selectAll("#pause").attr("fill", viewModel.settings.colorSelector.pauseColor.solid.color);
                this.svg.selectAll("#stop").attr("fill", viewModel.settings.colorSelector.stopColor.solid.color);
                this.svg.selectAll("#previous").attr("fill", viewModel.settings.colorSelector.previousColor.solid.color);
                this.svg.selectAll("#next").attr("fill", viewModel.settings.colorSelector.nextColor.solid.color);
            } else {
                let pickedColor = viewModel.settings.colorSelector.pickedColor.solid.color;
                this.svg.selectAll(".controls").attr("fill", viewModel.settings.colorSelector.pickedColor.solid.color);
            }
            let captionColor = viewModel.settings.captionSettings.captionColor.solid.color;      
            this.svg.select("#label").attr("fill", captionColor);

            //Change caption font size
            let fontSize = viewModel.settings.captionSettings.fontSize;
            this.svg.select("#label").attr("font-size", fontSize);

            let myViewBox = options.viewport;
            
            //Change title            
            if (this.visualSettings.captionSettings.show) {
                let title = options.dataViews[0].categorical.categories[0].source.displayName;           
                this.svg.select("#label").text(title);
                let textWidth = parseInt(this.svg.select("#label").text(title).style("width"));
                let viewBoxWidth = 155 + textWidth;
                this.controlsSVG
                .attr("viewBox","0 0 " + viewBoxWidth + " 24")
                .attr('preserveAspectRatio','xMinYMin');
                
                if (this.visualSettings.captionSettings.align == "right") {
                    this.captionSVG.select("text").attr('text-anchor', 'end').attr("x","100%");
                    this.captionSVG.attr("viewBox","0 0 145 24").attr('preserveAspectRatio','xMaxYMin');
                } else {
                    this.captionSVG.select("text").attr('text-anchor', 'start').attr("x","4%");
                    this.captionSVG.attr("viewBox","-140 0 " + viewBoxWidth + " 24").attr('preserveAspectRatio','xMinYMin');
                }
                
            } else {
                this.svg.select("#label").text("");
                this.controlsSVG
                .attr("viewBox","0 0 145 24")
                .attr('preserveAspectRatio','xMinYMin'); 
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

            //replay or stop after one cicle
            let stopAnimationTimer = setTimeout(() => {
                if(this.visualSettings.transitionSettings.loop) {
                    this.status = Status.Stop;
                    this.lastSelected = 0;
                    this.playAnimation();
                } else {
                    this.stopAnimation();
                }
            }, (this.viewModel.dataPoints.length - this.lastSelected) * timeInterval); 
            this.timers.push(stopAnimationTimer);
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
                            autoStart: this.visualSettings.transitionSettings.autoStart,
                            loop: this.visualSettings.transitionSettings.loop,
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
                    if (this.visualSettings.colorSelector.showAll) {
                        objectEnumeration.push({
                            objectName: objectName,
                            properties: {                                
                                showAll: this.visualSettings.colorSelector.showAll,
                                playColor: {
                                    solid: {
                                        color: this.visualSettings.colorSelector.playColor.solid.color
                                    }
                                },
                                pauseColor: {
                                   solid: {
                                       color: this.visualSettings.colorSelector.pauseColor.solid.color
                                   }
                                },
                                stopColor: {
                                   solid: {
                                       color: this.visualSettings.colorSelector.stopColor.solid.color
                                   }
                                },
                                previousColor: {
                                   solid: {
                                       color: this.visualSettings.colorSelector.previousColor.solid.color
                                   }
                                },
                                nextColor: {
                                    solid: {
                                       color: this.visualSettings.colorSelector.nextColor.solid.color
                                   }
                                }
                            },
                            selector: null
                        });
                    }  else {
                        objectEnumeration.push({
                        objectName: objectName,
                            properties: {
                                showAll: this.visualSettings.colorSelector.showAll,
                                pickedColor: {
                                    solid: {
                                        color: this.visualSettings.colorSelector.pickedColor.solid.color
                                    }
                                }
                            },
                            selector: null
                        });
                    }          
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
                            fontSize: this.visualSettings.captionSettings.fontSize,
                            align: this.visualSettings.captionSettings.align,
                        },
                        selector: null
                    });
                break;
            };
            return objectEnumeration;
        }
    }
}