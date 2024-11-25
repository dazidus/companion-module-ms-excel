module.exports = function (self) {
	self.setVariableDefinitions([
		{
			variableId: 'cell_value_A1',
			name: 'Value of Cell A1',
		},
		{
			variableId: 'cell_value_B1',
			name: 'Value of Cell B1',
		},
		{
			variableId: 'default_sheet_name',
			name: 'Default Sheet Name',
		},
	]);
};
async function updateVariables() {
	try {
		const valueA1 = await self.getCellValue('A1');
		const valueB1 = await self.getCellValue('B1');

		self.setVariableValues({
			cell_value_A1: valueA1,
			cell_value_B1: valueB1,
			default_sheet_name: self.config.sheetName,
		});
	} catch (error) {
		self.log('error', `Failed to update variables: ${error.message}`);
	}
}
