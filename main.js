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
			this.log('error', `Failed to authenticate with Microsoft Graph API: ${error.message}`);
			this.updateStatus(InstanceStatus.Error, 'Authentication failed');
		}

		this.updateActions();
		this.updateFeedbacks();
		this.updateVariableDefinitions();
	}

	async destroy() {
		this.log('debug', 'Module destroyed');
	}

	async configUpdated(config) {
		this.config = config;
		try {
			await this.verifyAPIAccess();
			this.updateStatus(InstanceStatus.Ok);
		} catch (error) {
			this.log('error', `Failed to reauthenticate with Microsoft Graph API: ${error.message}`);
			this.updateStatus(InstanceStatus.Error, 'Reauthentication failed');
		}
	}

	// Define configuration fields for the web interface
	getConfigFields() {
		return [
			{
				type: 'textinput',
				id: 'host',
				label: 'OneDrive/SharePoint URL',
				width: 8,
				required: true,
			},
			{
				type: 'textinput',
				id: 'clientId',
				label: 'Client ID',
				width: 6,
				required: true,
			},
			{
				type: 'textinput',
				id: 'clientSecret',
				label: 'Client Secret',
				width: 6,
				required: true,
			},
			{
				type: 'textinput',
				id: 'fileId',
				label: 'Excel File ID',
				width: 8,
				required: true,
			},
			{
				type: 'textinput',
				id: 'sheetName',
				label: 'Default Sheet Name',
				width: 4,
				required: true,
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
			throw new Error('No valid access token');
		}
	}

	async getAccessToken() {
		if (!this.config.clientId || !this.config.clientSecret) {
			this.log('error', 'Client ID or Client Secret is missing in the configuration');
			return null;
		}

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
			this.log('error', `Failed to retrieve access token: ${error.response?.data?.error_description || error.message}`);
			return null;
		}
	}

	async updateCellValue(cellAddress, value) {
		const token = await this.getAccessToken();
		if (!token) {
			this.log('error', 'Unable to update cell value: no access token');
			return;
		}

		if (!this.config.fileId || !this.config.sheetName) {
			this.log('error', 'File ID or Sheet Name is missing in the configuration');
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
			this.log('error', `Failed to update cell value: ${error.response?.data?.error?.message || error.message}`);
		}
	}
}

runEntrypoint(ModuleInstance, UpgradeScripts);
