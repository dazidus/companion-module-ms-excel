const { combineRgb } = require('@companion-module/base');

module.exports = async function (self) {
	self.setFeedbackDefinitions({
		cell_value_monitor: {
			name: 'Cell Value Monitor',
			type: 'boolean',
			label: 'Monitor Cell Value',
			defaultStyle: {
				bgcolor: combineRgb(0, 255, 0), // Grün für positive Bedingung
				color: combineRgb(0, 0, 0),
			},
			options: [
				{
					id: 'cell',
					type: 'textinput',
					label: 'Cell Address (e.g., A1)',
					default: 'A1',
				},
				{
					id: 'threshold',
					type: 'number',
					label: 'Threshold Value',
					default: 10,
					min: 0,
					max: 1000,
				},
			],
			callback: async (feedback) => {
				try {
					const currentValue = await self.getCellValue(feedback.options.cell);
					self.log('info', `Current value in ${feedback.options.cell}: ${currentValue}`);
					return parseFloat(currentValue) > feedback.options.threshold;
				} catch (error) {
					self.log('error', `Failed to retrieve cell value: ${error.message}`);
					return false;
				}
			},
		},
	});
};
