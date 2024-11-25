const { InstanceBase, Regex, runEntrypoint, InstanceStatus } = require('@companion-module/base');
const UpgradeScripts = require('./upgrades');
const UpdateActions = require('./actions');
const UpdateFeedbacks = require('./feedbacks');
const UpdateVariableDefinitions = require('./variables');
const axios = require('axios');

class ModuleInstance extends InstanceBase {
	constructor(internal) {
		super(internal);
	}

	async init(config) {
		this.config = config;

		// Verify API credentials and set the module status
		try {
			await this.verifyAPIAccess();
			this.updateStatus(InstanceStatus.Ok);
		} catch (error) {
			this.log('error', 'Failed to authenticate with Microsoft Graph API');
			this.updateStatus(InstanceStatus.Error);
		}

		this.updateActions(); // export actions
		this.updateFeedbacks(); // export feedbacks
		this.updateVariableDefinitions(); // export variable definitions
	}

	async destroy() {
		this.log('debug', 'destroy');
	}

	async configUpdated(config) {
		this.config = config;
		try {
			await this.verifyAPIAccess();
			this.updateStatus(InstanceStatus.Ok);
		} catch (error) {
			this.updateStatus(InstanceStatus.Error);
		}
	}

	// Return config fields for web config
	getConfigFields() {
		return [
			{
				type: 'textinput',
				id: 'host',
				label: 'OneDrive/SharePoint URL',
				width: 8,
			},
			{
				type: 'textinput',
				id: 'clientId',
				label: 'Client ID',
				width: 6,
			},
			{
				type: 'textinput',
				id: 'clientSecret',
				label: 'Client Secret',
				width: 6,
			},
			{
				type: 'textinput',
				id: 'fileId',
				label: 'Excel File ID',
				width: 8,
			},
			{
				type: 'textinput',
				id: 'sheetName',
				label: 'Default Sheet Name',
				width: 4,
			},
		];
	}

	updateActions() {
		UpdateActions(this);
	}

	updateFeedbacks() {
		UpdateFeedbacks(this);
	}

	updateVariableDefinitions() {
		UpdateVariableDefinitions(this);
	}

	async verifyAPIAccess() {
		const token = await this.getAccessToken();
		if (!token) {
			throw new Error('Authentication failed');
		}
	}

	async getAccessToken() {
		try {
			const response = await axios.post(
				'https://login.microsoftonline.com/common/oauth2/v2.0/token',
				new URLSearchParams({
					client_id: this.config.clientId,
					client_secret: this.config.clientSecret,
					grant_type: 'client_credentials',
					scope: 'https://graph.microsoft.com/.default',
				})
			);
			return response.data.access_token;
		} catch (error) {
			this.log('error', `Failed to retrieve access token: ${error.message}`);
			return null;
		}
	}

	async updateCellValue(cellAddress, value) {
		const token = await this.getAccessToken();
		if (!token) {
			this.log('error', 'Unable to update cell value: no access token');
			return;
		}

		try {
			const url = `https://graph.microsoft.com/v1.0/me/drive/items/${this.config.fileId}/workbook/worksheets/${this.config.sheetName}/range(address='${cellAddress}')`;
			await axios.patch(
				url,
				{ values: [[value]] },
				{ headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
			);
			this.log('info', `Cell ${cellAddress} updated to ${value}`);
		} catch (error) {
			this.log('error', `Failed to update cell value: ${error.message}`);
		}
	}
}

runEntrypoint(ModuleInstance, UpgradeScripts);
