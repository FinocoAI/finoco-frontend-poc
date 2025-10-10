/**
 * getChartSourceData.js
 *
 * Tool: Get the data source for an existing chart
 * Executor: Frontend (requires Office.js)
 */

export const toolDefinition = {
  name: "getChartSourceData",
  description: "Get the data source range and values for an existing chart",
  executor: "frontend",
  parameters: {
    sheetName: {
      type: "string",
      required: true,
      description: "Name of the sheet containing the chart"
    },
    chartName: {
      type: "string",
      required: true,
      description: "Name or title of the chart"
    }
  },
  returns: {
    sheetName: "string",
    chartName: "string",
    chartType: "string",
    sourceRange: "string",
    sourceData: "array[][]"
  }
};

export async function execute(params) {
  console.log(`🔧 Executing getChartSourceData:`, params);

  return Excel.run(async (context) => {
    try {
      const { sheetName, chartName } = params;

      if (!sheetName || typeof sheetName !== 'string') {
        throw new Error('Invalid sheetName');
      }
      if (!chartName || typeof chartName !== 'string') {
        throw new Error('Invalid chartName');
      }

      // Get worksheet
      const worksheet = context.workbook.worksheets.getItem(sheetName);

      // Try to get chart by name (this might be the internal name or title)
      // Office.js uses chart names, not titles
      let chart;
      try {
        chart = worksheet.charts.getItem(chartName);
      } catch (error) {
        // If not found by name, try finding by title
        const charts = worksheet.charts;
        charts.load("items");
        await context.sync();

        let foundChart = null;
        for (let i = 0; i < charts.items.length; i++) {
          const c = charts.items[i];
          c.load("name, title");
          await context.sync();

          if (c.title.text === chartName || c.name === chartName) {
            foundChart = c;
            break;
          }
        }

        if (!foundChart) {
          throw new Error(`Chart "${chartName}" not found in sheet "${sheetName}"`);
        }

        chart = foundChart;
      }

      // Load chart properties
      chart.load("name, chartType");
      await context.sync();

      // Get chart series to find data source
      const series = chart.series;
      series.load("items");
      await context.sync();

      if (series.items.length === 0) {
        throw new Error('Chart has no data series');
      }

      // Get first series data source
      const firstSeries = series.items[0];
      firstSeries.load("name");
      await context.sync();

      // Get the chart's data range
      // Note: This is a limitation - Office.js doesn't directly expose the full source range
      // We'll try to get it from the worksheet data source
      let sourceAddress = null;
      let sourceData = null;

      try {
        // Try to get the worksheet data source if available
        const dataSource = chart.getDataSourceString();
        dataSource.load();
        await context.sync();

        sourceAddress = dataSource.value;

        // Get data from that range
        const sourceRange = worksheet.getRange(sourceAddress);
        sourceRange.load("values");
        await context.sync();

        sourceData = sourceRange.values;
      } catch (error) {
        console.warn('  ⚠️ Could not retrieve full data source:', error.message);
        // Chart data source retrieval is limited in Office.js
        // We can still return chart metadata
      }

      const result = {
        sheetName: sheetName,
        chartName: chart.name,
        chartType: chart.chartType,
        sourceRange: sourceAddress || "Not available",
        sourceData: sourceData || []
      };

      console.log(`  ✅ Retrieved chart "${chartName}" (${result.chartType})`);

      return result;

    } catch (error) {
      console.error(`  ❌ Error in getChartSourceData:`, error);
      throw {
        tool: "getChartSourceData",
        error: error.message,
        params: params
      };
    }
  });
}
