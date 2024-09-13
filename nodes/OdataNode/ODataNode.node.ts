import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
	IDataObject,
} from 'n8n-workflow';
import { o } from 'odata';


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
		properties: [
			{
				displayName: 'URL',
				name: 'url',
				type: 'string',
				default: 'https://services.odata.org/TripPinRESTierService/',
				placeholder: 'https://services.odata.org/TripPinRESTierService',
				description: 'The OData service URL',
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
		],
	};


	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		let item: INodeExecutionData;
		let url: string;
		let method: string;
		let resource: string;
		let query: { [key: string]: any };
		let select: string;
		let filter: string;
		let orderby: string;
		let top: string;
		let skip: string;
		let data: { [key: string]: any };
		let response: IDataObject = {};

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				// OData service URL		
				method = this.getNodeParameter('method', itemIndex, '') as string;
				url = this.getNodeParameter('url', itemIndex, '') as string;
				resource = this.getNodeParameter('resource', itemIndex, '') as string;
				resource = resource.replace('"',"'")
				let query_str = this.getNodeParameter('query', itemIndex, '{}') as string;
				query = JSON.parse(query_str || '{}')
				select = this.getNodeParameter('select', itemIndex, '') as string;
				filter = this.getNodeParameter('filter', itemIndex, '') as string;
				orderby = this.getNodeParameter('orderby', itemIndex, '') as string;
				top = this.getNodeParameter('top', itemIndex, '') as string;
				skip = this.getNodeParameter('skip', itemIndex, '') as string;
				let data_str = this.getNodeParameter('data', itemIndex, '{}') as string;
				data = JSON.parse(data_str || '{}')
				
				//If no raw query given, build it from other fields
				if(!query || !Object.keys(query).length){
					query = {}
					if(select)
						query["$select"] =  select.replace(/\ /g, '')
					if(filter)
						query["$filter"] = filter
					if(orderby)
						query["$orderby"] = orderby
					if(top)
						query["$top"] = top
					if(skip)
						query["$skip"] = skip
				}

				console.log(method, resource, 'with query:', query, 'and data:', data)
				switch(method){
					case 'GET':
						response = await o(url)
							.get(resource)
							.query(query);
						break;
					case 'POST':
						response = await o(url)
							.post(resource, data)
							.query(query);
						break
					case 'PATCH':
						response = await o(url)
							.patch(resource, data)
							.query(query);
						break
					case 'DELETE':
						response = await o(url)
							.delete(resource)
							.query(query);
						break
				}

				item = items[itemIndex];
				item.json = response;

			} catch (error) {
				if (this.continueOnFail()) {
					items.push({ json: this.getInputData(itemIndex)[0].json, error, pairedItem: itemIndex });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						// If the error thrown already contains the context property,
						// only append the itemIndex
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
		}
	}
		return this.prepareOutputData(items);


	}
}
