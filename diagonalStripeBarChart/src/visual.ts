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
"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import DataViewCategoryColumn = powerbi.DataViewCategoryColumn;
import PrimitiveValue = powerbi.PrimitiveValue;
import IViewport = powerbi.IViewport;

import { VisualFormattingSettingsModel } from "./settings";

interface BarDataPoint {
    label: string;
    value: number;
    categoryDate: Date | null;
}

interface ChartData {
    isDateAxis: boolean;
    dataPoints: BarDataPoint[];
}

type StripeMode = "all" | "forecast";

export class Visual implements IVisual {
    private static readonly defaultBarColor: string = "#3a86ff";

    private readonly canvas: HTMLCanvasElement;
    private readonly context: CanvasRenderingContext2D;
    private formattingSettings: VisualFormattingSettingsModel;
    private readonly formattingSettingsService: FormattingSettingsService;
    private readonly valueFormatter: Intl.NumberFormat;

    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.valueFormatter = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });

        options.element.classList.add("diagonal-stripe-bar-chart");
        this.canvas = document.createElement("canvas");
        options.element.appendChild(this.canvas);

        const context = this.canvas.getContext("2d");
        if (!context) {
            throw new Error("Unable to create canvas rendering context.");
        }
        this.context = context;
    }

    public update(options: VisualUpdateOptions) {
        const dataView: DataView | undefined = options.dataViews?.[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);
        this.resizeCanvas(options.viewport);

        const chartData = this.buildChartData(dataView);
        if (!chartData || chartData.dataPoints.length === 0) {
            this.drawNoDataMessage(options.viewport);
            return;
        }

        this.drawChart(chartData, options.viewport);
    }

    /**
     * Returns properties pane formatting model content hierarchies, properties and latest formatting values, Then populate properties pane.
     * This method is called once every time we open properties pane or when the user edit any format property. 
     */
    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    private resizeCanvas(viewport: IViewport): void {
        const devicePixelRatio = window.devicePixelRatio || 1;
        const width = Math.max(Math.floor(viewport.width), 1);
        const height = Math.max(Math.floor(viewport.height), 1);

        this.canvas.style.width = `${width}px`;
        this.canvas.style.height = `${height}px`;
        this.canvas.width = Math.max(Math.floor(width * devicePixelRatio), 1);
        this.canvas.height = Math.max(Math.floor(height * devicePixelRatio), 1);
        this.context.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0);
        this.context.clearRect(0, 0, width, height);
    }

    private buildChartData(dataView: DataView | undefined): ChartData | null {
        const categorical = dataView?.categorical;
        const categoryColumn = categorical?.categories?.[0];
        const measureColumn = categorical?.values?.[0];

        if (!categoryColumn || !measureColumn) {
            return null;
        }

        const isDateAxis = this.isDateCategory(categoryColumn);
        const rowCount = Math.min(categoryColumn.values.length, measureColumn.values.length);
        const dataPoints: BarDataPoint[] = [];

        for (let index = 0; index < rowCount; index++) {
            const categoryValue = categoryColumn.values[index];
            const categoryDate = isDateAxis ? this.toDate(categoryValue, true) : null;
            const rawValue = measureColumn.values[index];
            const numericValue = this.toNumber(rawValue);

            dataPoints.push({
                label: this.getCategoryLabel(categoryValue, categoryDate),
                value: numericValue,
                categoryDate
            });
        }

        return {
            isDateAxis,
            dataPoints
        };
    }

    private drawNoDataMessage(viewport: IViewport): void {
        this.context.clearRect(0, 0, viewport.width, viewport.height);
        this.context.fillStyle = "#666";
        this.context.font = "14px sans-serif";
        this.context.textAlign = "center";
        this.context.textBaseline = "middle";
        this.context.fillText("Add fields to X axis and Y axis.", viewport.width / 2, viewport.height / 2);
    }

    private drawChart(chartData: ChartData, viewport: IViewport): void {
        const width = viewport.width;
        const height = viewport.height;
        const margin = { top: 20, right: 20, bottom: 72, left: 62 };
        const plotWidth = Math.max(width - margin.left - margin.right, 1);
        const plotHeight = Math.max(height - margin.top - margin.bottom, 1);

        const values = chartData.dataPoints.map((point) => point.value);
        const minValue = Math.min(0, ...values);
        const maxValue = Math.max(0, ...values);
        const valueRange = maxValue - minValue || 1;
        const yForValue = (value: number): number => margin.top + ((maxValue - value) / valueRange) * plotHeight;
        const baselineY = this.clamp(yForValue(0), margin.top, margin.top + plotHeight);
        const xAxisY = baselineY;

        this.context.clearRect(0, 0, width, height);
        this.drawYAxisGrid(minValue, maxValue, yForValue, margin.left, plotWidth);

        const stripeMode = this.getStripeMode();
        const barColor = this.getBarColor();
        const bandWidth = plotWidth / chartData.dataPoints.length;
        const barWidth = Math.max(Math.min(48, bandWidth * 0.7), 2);

        chartData.dataPoints.forEach((point, index) => {
            const centerX = margin.left + (index + 0.5) * bandWidth;
            const barX = centerX - barWidth / 2;
            const barEndY = yForValue(point.value);
            const barY = Math.min(baselineY, barEndY);
            const rawBarHeight = Math.abs(baselineY - barEndY);
            const barHeight = rawBarHeight < 1 && point.value !== 0 ? 1 : rawBarHeight;

            this.context.fillStyle = barColor;
            this.context.fillRect(barX, barY, barWidth, barHeight);

            if (barHeight > 0 && this.shouldDrawDiagonalStripes(stripeMode, chartData.isDateAxis, point.categoryDate)) {
                this.drawDiagonalStripes(barX, barY, barWidth, barHeight, barColor);
            }
        });

        this.drawAxes(margin.left, margin.top, margin.top + plotHeight, margin.left + plotWidth, xAxisY);
        this.drawCategoryLabels(chartData.dataPoints, margin.left, margin.top + plotHeight + 10, bandWidth, plotWidth);
    }

    private drawYAxisGrid(
        minValue: number,
        maxValue: number,
        yForValue: (value: number) => number,
        left: number,
        width: number
    ): void {
        const tickCount = 5;
        const ticks = this.getTicks(minValue, maxValue, tickCount);

        this.context.save();
        this.context.strokeStyle = "#e6e6e6";
        this.context.lineWidth = 1;
        this.context.font = "11px sans-serif";
        this.context.fillStyle = "#6a6a6a";
        this.context.textAlign = "right";
        this.context.textBaseline = "middle";

        ticks.forEach((tickValue) => {
            const y = yForValue(tickValue);
            this.context.beginPath();
            this.context.moveTo(left, y);
            this.context.lineTo(left + width, y);
            this.context.stroke();
            this.context.fillText(this.valueFormatter.format(tickValue), left - 8, y);
        });

        this.context.restore();
    }

    private drawAxes(left: number, top: number, bottom: number, right: number, xAxisY: number): void {
        this.context.save();
        this.context.strokeStyle = "#5a5a5a";
        this.context.lineWidth = 1;

        this.context.beginPath();
        this.context.moveTo(left, top);
        this.context.lineTo(left, bottom);
        this.context.stroke();

        this.context.beginPath();
        this.context.moveTo(left, xAxisY);
        this.context.lineTo(right, xAxisY);
        this.context.stroke();

        this.context.restore();
    }

    private drawCategoryLabels(
        dataPoints: BarDataPoint[],
        left: number,
        labelY: number,
        bandWidth: number,
        plotWidth: number
    ): void {
        const labelStep = Math.max(1, Math.ceil(dataPoints.length / Math.max(1, Math.floor(plotWidth / 70))));

        this.context.save();
        this.context.fillStyle = "#333";
        this.context.font = "11px sans-serif";
        this.context.textAlign = "center";
        this.context.textBaseline = "top";

        for (let index = 0; index < dataPoints.length; index += labelStep) {
            const label = this.truncateLabel(dataPoints[index].label, 12);
            const x = left + (index + 0.5) * bandWidth;
            this.context.fillText(label, x, labelY);
        }

        this.context.restore();
    }

    private drawDiagonalStripes(
        x: number,
        y: number,
        width: number,
        height: number,
        baseFillColor: string
    ): void {
        // Fixed 45-degree stripes as requested.
        const stripeSpacing = 8;
        this.context.save();
        this.context.beginPath();
        this.context.rect(x, y, width, height);
        this.context.clip();

        this.context.strokeStyle = this.getStripeColor(baseFillColor);
        this.context.lineWidth = 2;

        for (let startX = x - height; startX <= x + width + height; startX += stripeSpacing) {
            this.context.beginPath();
            this.context.moveTo(startX, y + height);
            this.context.lineTo(startX + height, y);
            this.context.stroke();
        }

        this.context.restore();
    }

    private shouldDrawDiagonalStripes(stripeMode: StripeMode, isDateAxis: boolean, categoryDate: Date | null): boolean {
        if (stripeMode === "all") {
            return true;
        }

        if (!isDateAxis || !categoryDate) {
            return false;
        }

        return this.isTodayOrFuture(categoryDate);
    }

    private getBarColor(): string {
        const configuredColor = this.formattingSettings?.appearanceCard?.barColor?.value?.value;
        if (configuredColor && typeof configuredColor === "string") {
            return configuredColor;
        }

        return Visual.defaultBarColor;
    }

    private getStripeMode(): StripeMode {
        const dropdownValue = this.formattingSettings?.appearanceCard?.stripeMode?.value;
        const selected = dropdownValue?.value;

        return selected === "forecast" ? "forecast" : "all";
    }

    private isDateCategory(categoryColumn: DataViewCategoryColumn): boolean {
        const categoryType = categoryColumn.source?.type as { dateTime?: boolean; temporal?: unknown } | undefined;
        if (categoryType?.dateTime || Boolean(categoryType?.temporal)) {
            return true;
        }

        return categoryColumn.values.some((value: PrimitiveValue) => value instanceof Date);
    }

    private toDate(value: PrimitiveValue, allowBroadParse: boolean = false): Date | null {
        if (value instanceof Date) {
            return Number.isNaN(value.getTime()) ? null : new Date(value.getTime());
        }

        if (typeof value === "string") {
            if (!allowBroadParse && !this.looksLikeDateString(value)) {
                return null;
            }

            const parsed = new Date(value);
            return Number.isNaN(parsed.getTime()) ? null : parsed;
        }

        if (allowBroadParse && typeof value === "number") {
            const parsed = new Date(value);
            return Number.isNaN(parsed.getTime()) ? null : parsed;
        }

        return null;
    }

    private looksLikeDateString(value: string): boolean {
        const trimmed = value.trim();
        return /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(trimmed);
    }

    private isTodayOrFuture(dateValue: Date): boolean {
        const dayStart = new Date(dateValue.getTime());
        dayStart.setHours(0, 0, 0, 0);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return dayStart.getTime() >= today.getTime();
    }

    private toNumber(value: PrimitiveValue): number {
        if (typeof value === "number" && Number.isFinite(value)) {
            return value;
        }

        if (value === null || value === undefined) {
            return 0;
        }

        const numeric = Number(value);
        return Number.isFinite(numeric) ? numeric : 0;
    }

    private getCategoryLabel(categoryValue: PrimitiveValue, categoryDate: Date | null): string {
        if (categoryDate) {
            return categoryDate.toLocaleDateString(undefined, {
                year: "numeric",
                month: "short",
                day: "numeric"
            });
        }

        if (categoryValue === null || categoryValue === undefined) {
            return "(blank)";
        }

        return String(categoryValue);
    }

    private getTicks(minValue: number, maxValue: number, count: number): number[] {
        if (count <= 1) {
            return [minValue, maxValue];
        }

        const range = maxValue - minValue;
        if (range === 0) {
            return [minValue];
        }

        const tickValues: number[] = [];
        const step = range / (count - 1);
        for (let index = 0; index < count; index++) {
            tickValues.push(minValue + step * index);
        }

        return tickValues;
    }

    private truncateLabel(value: string, maxLength: number): string {
        if (value.length <= maxLength) {
            return value;
        }

        return `${value.slice(0, maxLength - 1)}\u2026`;
    }

    private getStripeColor(fillColor: string): string {
        const rgb = this.parseColor(fillColor);
        if (!rgb) {
            return "rgba(255, 255, 255, 0.65)";
        }

        const luminance = (0.299 * rgb.r + 0.587 * rgb.g + 0.114 * rgb.b) / 255;
        return luminance > 0.6 ? "rgba(0, 0, 0, 0.30)" : "rgba(255, 255, 255, 0.70)";
    }

    private parseColor(color: string): { r: number; g: number; b: number } | null {
        const hex = color.trim().toLowerCase();
        const shortHexMatch = /^#([0-9a-f]{3})$/.exec(hex);
        if (shortHexMatch) {
            const [r, g, b] = shortHexMatch[1].split("");
            return {
                r: parseInt(`${r}${r}`, 16),
                g: parseInt(`${g}${g}`, 16),
                b: parseInt(`${b}${b}`, 16)
            };
        }

        const longHexMatch = /^#([0-9a-f]{6})$/.exec(hex);
        if (longHexMatch) {
            const value = longHexMatch[1];
            return {
                r: parseInt(value.slice(0, 2), 16),
                g: parseInt(value.slice(2, 4), 16),
                b: parseInt(value.slice(4, 6), 16)
            };
        }

        const rgbMatch = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/.exec(hex);
        if (rgbMatch) {
            return {
                r: Number(rgbMatch[1]),
                g: Number(rgbMatch[2]),
                b: Number(rgbMatch[3])
            };
        }

        return null;
    }

    private clamp(value: number, minValue: number, maxValue: number): number {
        return Math.min(Math.max(value, minValue), maxValue);
    }
}