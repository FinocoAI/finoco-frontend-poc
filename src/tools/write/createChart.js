/**
 * createChart.js
 * Tool: Create a new chart
 *
 * Valid chartType values:
 * - ColumnClustered: Vertical column chart
 * - ColumnStacked: Stacked vertical columns
 * - BarClustered: Horizontal bar chart
 * - BarStacked: Stacked horizontal bars
 * - Line: Line chart
 * - LineMarkers: Line chart with markers
 * - Pie: Pie chart
 * - XYScatter: Scatter plot
 * - Area: Area chart
 * - AreaStacked: Stacked area chart
 */

export const toolDefinition = {
  name: "createChart",
  description: "Create a new chart in Excel",
  executor: "frontend",
  parameters: {
    sheetName: { type: "string", required: true },
    chartType: { type: "string", required: true },
    dataRange: { type: "string", required: true },
    title: { type: "string", required: false },
    position: { type: "string", required: false }
  }
};

export async function execute(params) {
  return Excel.run(async (context) => {
    const { sheetName, chartType, dataRange, title, position } = params;

    // Validate and normalize chart type
    const validChartTypes = {
      'bar': 'BarClustered',
      'barclustered': 'BarClustered',
      'barstacked': 'BarStacked',
      'column': 'ColumnClustered',
      'columnclustered': 'ColumnClustered',
      'columnstacked': 'ColumnStacked',
      'line': 'Line',
      'linemarkers': 'LineMarkers',
      'pie': 'Pie',
      'scatter': 'XYScatter',
      'xyscatter': 'XYScatter',
      'area': 'Area',
      'areastacked': 'AreaStacked'
    };

    const normalizedType = validChartTypes[chartType.toLowerCase()] || chartType;

    const worksheet = context.workbook.worksheets.getItem(sheetName);
    const sourceData = worksheet.getRange(dataRange);

    const chart = worksheet.charts.add(normalizedType, sourceData, "Columns");
    if (title) chart.title.text = title;

    chart.height = 300;
    chart.width = 500;

    if (position) {
      chart.setPosition(position);
    }

    await context.sync();

    return { sheetName, chartType: normalizedType, title: title || normalizedType };
  });
}
