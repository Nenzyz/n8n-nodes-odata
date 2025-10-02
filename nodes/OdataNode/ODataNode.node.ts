import {
	BINARY_ENCODING,
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	IRequestOptionsSimplified,
	jsonParse,
	NodeOperationError,
	IDataObject,
	IHttpRequestOptions
} from 'n8n-workflow';
import { o } from 'odata';
import type { Readable } from 'stream';

export async function reduceAsync<T, R>(
	arr: T[],
	reducer: (acc: Awaited<Promise<R>>, cur: T) => Promise<R>,
	init: Promise<R> = Promise.resolve({} as R),
): Promise<R> {
	return await arr.reduce(async (promiseAcc, item) => {
		return await reducer(await promiseAcc, item);
	}, init);
}

export const keysToLowercase = <T>(headers: T) => {
	if (typeof headers !== 'object' || Array.isArray(headers) || headers === null) return headers;
	return Object.entries(headers).reduce((acc, [key, value]) => {
		acc[key.toLowerCase()] = value as IDataObject;
		return acc;
	}, {} as IDataObject);
};

export class ODataNode implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'OData Node',
		name: 'odataNode',
		icon: 'file:odata.svg',
		group: ['transform'],
		version: 1,
		description: 'Interact with OData REST interface',
		defaults: {
			name: 'OData Node',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'oAuth2Api',
				required: false,
			},
			{
				name: 'oDataOAuth2Api',
				required: false,
			},
			{
				name: 'httpCustomAuth',
				required: false,
			}
		],
		properties: [
			{
				displayName: 'URL',
				name: 'url',
				type: 'string',
				default: 'https://graph.microsoft.com/v1.0/',
				placeholder: 'https://graph.microsoft.com/v1.0/',
				description: 'The OData service URL',
			},

			{
				displayName: 'Authentication',
				name: 'authentication',
				noDataExpression: true,
				type: 'options',
				options: [
					{
						name: 'None',
						value: 'none',
					},
					{
						name: 'Generic Credential Type',
						value: 'genericCredentialType',
						description: 'Fully customizable. Choose between basic, header, OAuth2, etc.',
					},
				],
				default: 'none',
			},
			{
				displayName: 'Credential Type',
				name: 'nodeCredentialType',
				type: 'credentialsSelect',
				noDataExpression: true,
				required: true,
				default: '',
				credentialTypes: ['extends:oAuth2Api', 'extends:oAuth1Api', 'has:authenticate'],
				displayOptions: {
					show: {
						authentication: ['predefinedCredentialType'],
					},
				},
			},
			{
				displayName: 'Generic Auth Type',
				name: 'genericAuthType',
				type: 'credentialsSelect',
				required: true,
				default: '',
				credentialTypes: ['has:genericAuth'],
				displayOptions: {
					show: {
						authentication: ['genericCredentialType'],
					},
				},
			},

			{
				displayName: "Method",
				name: "method",
				type: "options",
				options: [
					{
						"name": "GET",
						"value": "GET"
					},
					{
						"name": "POST",
						"value": "POST"
					},
					{
						"name": "PATCH",
						"value": "PATCH"
					},
					{
						"name": "DELETE",
						"value": "DELETE"
					}
				],
				default: "GET"
			},
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'string',
				default: "",
				placeholder: "People('scottketchum')",
				description: 'The OData resource to fetch',
			},
			{
				displayName: 'Data',
				name: 'data',
				type: 'string',
				default: '',
				placeholder: '{ "UserName": "newuser", "FirstName": "New", "LastName": "User" }',
				description: 'Data to POST as valid JSON string.',
				displayOptions: {
                    show: {
					   method: ['POST', 'PATCH']
                    },
                },
			},
			{
				displayName: 'Send Headers',
				name: 'sendHeaders',
				type: 'boolean',
				default: false,
				noDataExpression: true,
				description: 'Whether the request has headers or not',
			},
			{
				displayName: 'Specify Headers',
				name: 'specifyHeaders',
				type: 'options',
				displayOptions: {
					show: {
						sendHeaders: [true],
					},
				},
				options: [
					{
						name: 'Using Fields Below',
						value: 'keypair',
					},
					{
						name: 'Using JSON',
						value: 'json',
					},
				],
				default: 'keypair',
			},
			{
				displayName: 'Header Parameters',
				name: 'headerParameters',
				type: 'fixedCollection',
				displayOptions: {
					show: {
						sendHeaders: [true],
						specifyHeaders: ['keypair'],
					},
				},
				typeOptions: {
					multipleValues: true,
				},
				placeholder: 'Add Parameter',
				default: {
					parameters: [
						{
							name: '',
							value: '',
						},
					],
				},
				options: [
					{
						name: 'parameters',
						displayName: 'Parameter',
						values: [
							{
								displayName: 'Name',
								name: 'name',
								type: 'string',
								default: '',
							},
							{
								displayName: 'Value',
								name: 'value',
								type: 'string',
								default: '',
							},
						],
					},
				],
			},
			{
				displayName: 'JSON',
				name: 'jsonHeaders',
				type: 'json',
				displayOptions: {
					show: {
						sendHeaders: [true],
						specifyHeaders: ['json'],
					},
				},
				default: '',
			},
            {
                displayName: 'Advanced',
                name: 'visibleOption',
                type: 'boolean',
                default: false,
                description: 'Advanced Options',
            },
			{
				displayName: 'Raw Query',
				name: 'query',
				type: 'string',
				default: '',
				placeholder: `{"$filter": "FirstName eq 'John'", "$select": "FirstName,LastName"}`,
				description: 'The raw OData query, as valid JSON. Overrides other options.',
                displayOptions: {
                    show: {
                        visibleOption: [true],  // Show this option only if visibleOption is true
                    },
                },
			},
			{
				displayName: '$select',
				name: 'select',
				type: 'string',
				default: '',
				placeholder: 'FirstName,LastName,UserName',
				description: 'The fields to select, separated by commas',
				displayOptions: {
                    show: {
                       query: [''], //raw query overrides these controls
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$filter',
				name: 'filter',
				type: 'string',
				default: '',
				placeholder: "LastName eq 'Russell' or FirstName eq 'Scott'",
				description: 'The filter expression',
				displayOptions: {
                    show: {
                       query: [''],
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$expand',
				name: 'expand',
				type: 'string',
				default: '',
				placeholder: 'Products',
				description: 'The command to expand',
				displayOptions: {
                    show: {
                       query: [''], //raw query overrides these controls
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$orderby',
				name: 'orderby',
				type: 'string',
				default: '',
				placeholder: "LastName desc",
				description: 'The field to order by',
				displayOptions: {
                    show: {
                       query: [''],
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$top',
				name: 'top',
				type: 'number',
				default: '0',
				placeholder: '10',
				description: 'How many results to return',
				displayOptions: {
                    show: {
                       query: [''],
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$skip',
				name: 'skip',
				type: 'number',
				default: '',
				placeholder: '3',
				description: 'How many results to skip',
				displayOptions: {
                    show: {
                       query: [''],
					   visibleOption: [true]
                    },
                },
			},
			{
				displayName: '$count',
				name: 'count',
				type: 'boolean',
				default: false,
				description: 'Return inline count',
				displayOptions: {
                    show: {
                       query: [''],
					   visibleOption: [true]
                    },
                },
			},
		],
	};


	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		//let item: INodeExecutionData;
		let url: string;
		let method: string;
		let resource: string;
		let query: { [key: string]: any };
		let select: string;
		let filter: string;
		let orderby: string;
		let top: string;
		let skip: string;
		let expand: string;
		let count: boolean;
		let data: { [key: string]: any };
		let response: IDataObject[] = [];
		let authentication;



		let options: IHttpRequestOptions = {url:'', method:'GET'};

		let newitems: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			//Authentication
			try {
				authentication = this.getNodeParameter('authentication', 0) as
					//| 'predefinedCredentialType'
					| 'genericCredentialType'
					| 'none';
			} catch {}

			var oAuth2Api;
			var basicAuth;
			let httpCustomAuth;
			if (authentication === 'genericCredentialType') {
				let genericCredentialType = this.getNodeParameter('genericAuthType', 0) as string;


				console.log('genericCredentialType', genericCredentialType);

				if (genericCredentialType === 'oAuth2Api') {
					oAuth2Api = await this.getCredentials('oAuth2Api', itemIndex);
				} else if (genericCredentialType === 'basicAuth') {
					basicAuth = await this.getCredentials('basicAuth', itemIndex);
					console.log('basicAuth', basicAuth);
				} else if (genericCredentialType === 'httpCustomAuth') {
					httpCustomAuth = await this.getCredentials('httpCustomAuth', itemIndex);
					console.log('httpCustomAuth', httpCustomAuth);
				}
			}

			try {
				// OData service URL		
				method = this.getNodeParameter('method', itemIndex, '') as string;
				url = this.getNodeParameter('url', itemIndex, '') as string;
				if (url.slice(-1) !== '/') //Odata requires resource to end in /
					url += '/';

				resource = this.getNodeParameter('resource', itemIndex, '') as string;
				resource = resource.replace('"',"'")

				let query_str = this.getNodeParameter('query', itemIndex, '{}') as string;
				query = JSON.parse(query_str || '{}')
				select = this.getNodeParameter('select', itemIndex, '') as string;
				filter = this.getNodeParameter('filter', itemIndex, '') as string;
				orderby = this.getNodeParameter('orderby', itemIndex, '') as string;
				top = this.getNodeParameter('top', itemIndex, '') as string;
				skip = this.getNodeParameter('skip', itemIndex, '') as string;
				expand = this.getNodeParameter('expand', itemIndex, '') as string;
				count = this.getNodeParameter('count', itemIndex, '') as boolean;
				let data_str = this.getNodeParameter('data', itemIndex, '{}') as string;
				data = JSON.parse(data_str || '{}')
				options.url = url;
				options.method = 'GET';

				var customHeaders = {headers: {}};

				if (oAuth2Api?.oauthTokenData){
					let tokendata;
					if (typeof oAuth2Api?.oauthTokenData == 'string')
						tokendata = JSON.parse(oAuth2Api?.oauthTokenData)
					else
						tokendata = oAuth2Api?.oauthTokenData

					customHeaders = {headers:{'Authorization':`Bearer ${tokendata.access_token}`}}
				} else if (basicAuth) {
					const auth = new Buffer(`${basicAuth.user}:${basicAuth.password}`, 'binary').toString('base64');
					console.log(auth);
					customHeaders = {headers:{'Authorization':`Basic ${auth}`}}
				} else if (httpCustomAuth !== undefined){
					const customAuth = jsonParse<IRequestOptionsSimplified>(
						(httpCustomAuth.json as string) || '{}',
						{ errorMessage: 'Invalid Custom Auth JSON' },
					);
					console.log('customAuth', customAuth);
					if (customAuth.headers) {
						customHeaders = { ...customHeaders, ...customAuth };
					}
					console.log('customHeaders', customHeaders);
				}

				const parametersToKeyValue = async (
					accumulator: { [key: string]: any },
					cur: { name: string; value: string; parameterType?: string; inputDataFieldName?: string },
				) => {
					if (cur.parameterType === 'formBinaryData') {
						if (!cur.inputDataFieldName) return accumulator;
						const binaryData = this.helpers.assertBinaryData(itemIndex, cur.inputDataFieldName);
						let uploadData: Buffer | Readable;
						const itemBinaryData = items[itemIndex].binary![cur.inputDataFieldName];
						if (itemBinaryData.id) {
							uploadData = await this.helpers.getBinaryStream(itemBinaryData.id);
						} else {
							uploadData = Buffer.from(itemBinaryData.data, BINARY_ENCODING);
						}
	
						accumulator[cur.name] = {
							value: uploadData,
							options: {
								filename: binaryData.fileName,
								contentType: binaryData.mimeType,
							},
						};
						return accumulator;
					}
					accumulator[cur.name] = cur.value;
					return accumulator;
				};

				const sendHeaders = this.getNodeParameter('sendHeaders', itemIndex, false) as boolean;

				const headerParameters = this.getNodeParameter(
					'headerParameters.parameters',
					itemIndex,
					[],
				) as [{ name: string; value: string }];
	
				const specifyHeaders = this.getNodeParameter(
					'specifyHeaders',
					itemIndex,
					'keypair',
				) as string;
	
				const jsonHeadersParameter = this.getNodeParameter('jsonHeaders', itemIndex, '') as string;

				// Get parameters defined in the UI
				if (sendHeaders && headerParameters) {
					let additionalHeaders: IDataObject = {};
					if (specifyHeaders === 'keypair') {
						additionalHeaders = await reduceAsync(
							headerParameters.filter((header) => header.name),
							parametersToKeyValue,
						);
					} else if (specifyHeaders === 'json') {
						// body is specified using JSON
						try {
							JSON.parse(jsonHeadersParameter);
						} catch {
							throw new NodeOperationError(
								this.getNode(),
								'JSON parameter need to be an valid JSON',
								{
									itemIndex,
								},
							);
						}

						additionalHeaders = jsonParse(jsonHeadersParameter);
					}

					customHeaders.headers = {
						...(customHeaders.headers || {}),
						...keysToLowercase(additionalHeaders),
					};
				}

				let ohandler =  o(url, customHeaders)

				//If no raw query given, build it from other fields
				if(!query || !Object.keys(query).length){
					query = {}
					if(select)
						query["$select"] =  select.replace(/\ /g, '')
					if(filter)
						query["$filter"] = filter
					if(orderby)
						query["$orderby"] = orderby
					if(expand)
						query["$expand"] =  expand.replace(/\ /g, '')
					if(top)
						query["$top"] = top
					if(skip)
						query["$skip"] = skip
					if(count != undefined)
						query["$count"] = count
				}


				switch(method){
					case 'GET':
						if (count) {
							const buildUrl = url.endsWith('/')
								? `${url}${resource}`
								: `${url}/${resource}`;
							const fetchResponse = await fetchData(buildUrl, query, customHeaders.headers);
							response = fetchResponse;
							break;
						}
						response = await ohandler
							.get(resource)
							.query(query);
						break;
					case 'POST':
						response = await ohandler
							.post(resource, data)
							.query(query);
						break
					case 'PATCH':
						response = await ohandler
							.patch(resource, data)
							.query(query);
						break
					case 'DELETE':
						response = await ohandler
							.delete(resource)
							.query(query);
						break
				}


				if (!Array.isArray(response)) {
					response = [response]
				}

				for (let obj of response) {
					newitems.push({
						json: obj,
						pairedItem: { item: itemIndex, input: undefined }
					});
				  }

			} catch (error) {
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						error.context.itemIndex = itemIndex;
						throw error;
					}

					// Extract meaningful error information
					let errorMessage = 'An error occurred while executing the OData request';

					if (error instanceof Error) {
						errorMessage = error.message;
					} else if (typeof error === 'string') {
						errorMessage = error;
					} else if (error?.message) {
						errorMessage = error.message;
					}

					// Add additional context if available
					if (error?.response) {
						const status = error.response.status || error.response.statusCode;
						const statusText = error.response.statusText || error.response.statusMessage;
						if (status) {
							errorMessage = `HTTP ${status}${statusText ? ` ${statusText}` : ''}: ${errorMessage}`;
						}
						if (error.response.data) {
							try {
								const responseData = typeof error.response.data === 'string'
									? error.response.data
									: JSON.stringify(error.response.data, null, 2);
								errorMessage += `\n\nResponse: ${responseData}`;
							} catch {}
						}
					}

					throw new NodeOperationError(this.getNode(), errorMessage, {
						itemIndex,
						description: error?.stack || undefined,
					});
				}
			}
	}



		return this.prepareOutputData(newitems);

	}
}

const fetchData = async function(url: string, query: { [key: string]: any }, headers: any): Promise<IDataObject[]> {
    const fullUrl = `${url}?${new URLSearchParams(query).toString()}`;
	console.log('fullUrl', fullUrl);
    const response = await fetch(fullUrl, {
        method: "GET",
        headers: {
            "Accept": "application/json",
            "Content-Type": "application/json",
            ...headers,
        },
    });

    if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    if (!Array.isArray(data)) {
        return [data] as IDataObject[]; // Ensure it's always an array
    }

    return data as IDataObject[];
}