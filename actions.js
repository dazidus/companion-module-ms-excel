module.exports = function (self) {
	self.setActionDefinitions({
		set_cell_value: {
			name: 'Set Cell Value',
			options: [
				{
					id: 'cell',
					type: 'textinput',
					label: 'Cell Address (e.g., A1)',
					default: 'A1',
				},
				{
					id: 'value',
					type: 'textinput',
					label: 'Value to Set',
					default: '',
				},
			],
			callback: async (event) => {
				try {
					await self.updateCellValue(event.options.cell, event.options.value);
					self.log('info', `Cell ${event.options.cell} updated to ${event.options.value}`);
				} catch (error) {
					self.log('error', `Failed to set cell value: ${error.message}`);
				}
			},
		},
		increase_cell_value: {
			name: 'Increase Cell Value',
			options: [
				{
					id: 'cell',
					type: 'textinput',
					label: 'Cell Address (e.g., A1)',
					default: 'A1',
				},
				{
					id: 'increment',
					type: 'number',
					label: 'Increment Value',
					default: 1,
					min: 1,
					max: 1000,
				},
			],
			callback: async (event) => {
				try {
					const currentValue = await self.getCellValue(event.options.cell);
					const newValue = parseFloat(currentValue) + event.options.increment;
					await self.updateCellValue(event.options.cell, newValue);
					self.log('info', `Cell ${event.options.cell} increased to ${newValue}`);
				} catch (error) {
					self.log('error', `Failed to increase cell value: ${error.message}`);
				}
			},
		},
		decrease_cell_value: {
			name: 'Decrease Cell Value',
			options: [
				{
					id: 'cell',
					type: 'textinput',
					label: 'Cell Address (e.g., A1)',
					default: 'A1',
				},
				{
					id: 'decrement',
					type: 'number',
					label: 'Decrement Value',
					default: 1,
					min: 1,
					max: 1000,
				},
			],
			callback: async (event) => {
				try {
					const currentValue = await self.getCellValue(event.options.cell);
					const newValue = parseFloat(currentValue) - event.options.decrement;
					await self.updateCellValue(event.options.cell, newValue);
					self.log('info', `Cell ${event.options.cell} decreased to ${newValue}`);
				} catch (error) {
					self.log('error', `Failed to decrease cell value: ${error.message}`);
				}
			},
		},
	});
};
