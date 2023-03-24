// ShipPrint.cs
// 
// This program will ship a package list, display the results, and print package labels

// #define DEVICE_AVAILABLE

using System;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using Progistics.API;
using Progistics.Base;
using Progistics.TransAPI;
using Progistics.Devices;
using F = Progistics.TransAPI.TransAPIDictionary.Fields;

#pragma warning disable 0162 // Ignore warnings about unreachable code

namespace ShipPrint
{
	public class ShipPrint
	{
		// Default shipper
		const string shipperSymbol = "TEX";

		// Package ID
		const string packageID = "1020";

		// Set to 'false' to disable document generation
		const bool generateDocuments = true;

		// Set to 'true' to closeout existing shipments and don't do anything else
		const bool closeOutOnly = false;

		// Use current date for ship date
		static readonly ITDCDate shipDate = new CoDate();

		// Database connection string
		const string sqlConnectionString = @"Data Source=.\TRAINING;Initial Catalog=OrdersDB;User=sa;Password=C0nnectsh!p";

		// Database query for orders
		const string sqlSelectStatement = "SELECT * FROM [orders] WHERE [db_order] = '" + packageID + "'";
		//const string sqlSelectStatement = "SELECT TOP 1 * FROM [orders] WHERE [db_company] LIKE '%wal-mart%' ORDER BY [db_order] ASC";

		// query strings for Global lab, 1041, 1042
		//const string sqlSelectStatement = "SELECT TOP 1 * FROM [orders] where db_order = 1040 ORDER BY [db_order] ASC";

		// Document output
		public const DocumentOutput documentOutput = DocumentOutput.Pdf;

		private static void Main()
		{
			// Create instance of TransAPI
			Console.WriteLine("Initializing TransAPI");
			ITransAPI transAPI = new CoTransAPI();

			try
			{
				// Initialize transAPI's internal information/lists
				ITDCReturn ret = transAPI.Init(false);
				if (!ret.IsSuccess)
				{
					Console.WriteLine("TransAPI Initialization Error: {0} ({1}).", ret.Message, ret.Code);
					return;
				}

				if (closeOutOnly)
				{
					performCloseout(transAPI, "CONNECTSHIP_UPS.UPS", generateDocuments);
					return;
				}

				// Get orders from database
				using (SqlConnection dbConn = new SqlConnection(sqlConnectionString))
				using (SqlCommand cmd = new SqlCommand(sqlSelectStatement, dbConn))
				{
					dbConn.Open();
					using (SqlDataReader reader = cmd.ExecuteReader())
					{
						while (reader.Read())
						{
							// ---- Prepare package data ----
							Console.WriteLine("\nPreparing package data");

							// Map database country to Progistics country symbol
							string progisticsCountrySymbol = getCountrySymbol(Convert.ToString(reader["db_country"]));
							if (progisticsCountrySymbol == null)
							{
								Console.Error.WriteLine("*** Error: Unable to map country {0} to Progistics country symbol ***", Convert.ToString(reader["db_country"]));
								continue;
							}

							// Map database service to Progistics service symbol
							string progisticsServiceSymbol = getProgisticsService(Convert.ToString(reader["db_service"]));
							if (progisticsServiceSymbol == null)
							{
								Console.Error.WriteLine("*** Error: Unable to map service {0} to Progistics service symbol ***", Convert.ToString(reader["db_service"]));
								continue;
							}

							// Consignee from database
							CoNameAddress consignee = new CoNameAddress();
							consignee.CountrySymbol = progisticsCountrySymbol;
							consignee.Company = Convert.ToString(reader["db_company"]);
							consignee.Contact = Convert.ToString(reader["db_attn"]);
							consignee.Address1 = Convert.ToString(reader["db_street"]);
							consignee.Address2 = Convert.ToString(reader["db_room"]);
							consignee.Address3 = Environment.MachineName;
							consignee.Phone = Convert.ToString(reader["db_phone"]);
							consignee.City = Convert.ToString(reader["db_city"]);
							consignee.StateProvince = Convert.ToString(reader["db_state"]);
							consignee.PostalCode = Convert.ToString(reader["db_postal"]);
							consignee.Residential = Convert.ToBoolean(reader["db_residential"]);

							// Populate the default atributes dictionary
							CoDictionary defaultAttributes = new CoDictionary();
							defaultAttributes[F.CONSIGNEE] = consignee;
							defaultAttributes[F.SHIPPER] = shipperSymbol;
							defaultAttributes[F.SHIPDATE] = shipDate;

							// Create list to contain packages
							CoSimpleList packages = new CoSimpleList();

							// Create package from orders database
							CoDictionary package = new CoDictionary();
							package[F.WEIGHT] = reader["db_weight"];
							package[F.DIMENSION] = reader["db_dimension"];
							package[F.PACKAGING] = reader["db_packaging"];
							package[F.TERMS] = reader["db_terms"];
							package[F.DESCRIPTION] = reader["db_description"];

							package[F.SHIPPER_REFERENCE] = "ABC-12345";
							package[F.CONSIGNEE_REFERENCE] = "P.O. 99999999";

							if (consignee.CountrySymbol != "UNITED_STATES")
							{
								addInternationalAttributes(package);
							}

							if (consignee.Company.ToLower().Contains("wal-mart"))
							{
								CoDictionary product = new CoDictionary();
								product["pronum"] = "0123456789";
								product["blnum"] = "0123456789";
								product["wmv"] = "123456789";
								product["storenum"] = "00656";
								product["loc"] = "06094";
								product["type"] = "0073";
								product["dept"] = "0052";
								product["ordernum"] = "1234567890";
								package[F.USER_DATA_1] = product;
							}

							// Add package to the packages list
							string orderId = Convert.ToString(reader["db_order"]);
							packages.Add(package);

							// add an additional package for order 1041
							if (orderId.Equals("1041"))
							{
								CoDictionary package2 = new CoDictionary();
								package2[F.WEIGHT] = Convert.ToDouble(reader["db_weight"]) - 15;
								package2[F.DIMENSION] = "12x12x12";
								package2[F.PACKAGING] = reader["db_packaging"];
								package2[F.TERMS] = reader["db_terms"];
								package2[F.DESCRIPTION] = reader["db_description"];

								package2[F.SHIPPER_REFERENCE] = "ABC-12345";
								package2[F.CONSIGNEE_REFERENCE] = "P.O. 99999999";
								packages.Add(package2);
							}

							Console.WriteLine("Package for order {0}:", orderId);
							displayDictionaryData(package);

							// Display default attributes
							Console.WriteLine("Default package attributes:");
							displayDictionaryData(defaultAttributes);

							// Ship package
							Console.WriteLine("\nCalling Ship");
							ITransAPIResult result = transAPI.Ship(defaultAttributes, packages, progisticsServiceSymbol, CloseOutMode.comRelease);

							// Display shipping results
							Console.WriteLine("\nDisplaying Shipping Results");

							Console.WriteLine("Service: {0}", result.Subcategory.FriendlyName);
							Console.WriteLine("  Shipment Result: {0} ({1})", result.ErrorMessage, result.ErrorCode);

							// Get the result data for the package list
							ITransAPIData shipmentResultData = result.ResultData;

							// Display shipment-level data
							displayTransAPIData(shipmentResultData);

							// Display the rating or shipping results for each package
							int packageCount = 0;

							CoSimpleList bundleIdList = new CoSimpleList();

							foreach (ITransAPIPackageResult packageResult in result.PackageResultList)
							{
								Console.WriteLine("Package {0}:", ++packageCount);
								Console.WriteLine("  Result: {0} ({1})", packageResult.ErrorMessage, packageResult.ErrorCode);

								// Get the result data for the package
								ITransAPIData packageResultData = packageResult.ResultData;

								// Display package data
								displayTransAPIData(packageResultData);

								// Print document and save order information to update database
								if (packageResult.ErrorCode == (int)ErrorCode.ecNoError)
								{
									if (generateDocuments)
										generateAndPrintDocument(transAPI, (int)packageResultData[F.MSN], (int)packageResultData[F.BUNDLE_ID], progisticsServiceSymbol, shipperSymbol);

									// Update order in database
									updateDB(orderId, packageResult);
								}
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("   Error: {0}", ex.Message);
			}
			finally
			{
				// Cleanup
				transAPI.DestroyNotify();
			}
		}

		private static void performCloseout(ITransAPI transAPI, string category, bool generateDocuments)
		{
			Console.WriteLine("Closing out for category {0}, shipper {1}", category, shipperSymbol);

			// Enumerate open manifests
			ITDCReturn ret = transAPI.GetCloseOutItemList(category, shipperSymbol);
			if (!ret.IsSuccess)
			{
				Console.WriteLine("GetCloseOutItemList returned: {0} ({1})", ret.Message, ret.Code);
				return;
			}

			// Get the list of manifests
			ITDCCollection list = (ITDCCollection)ret.Data;
			if (list.Count == 0)
			{
				Console.WriteLine("There are no manifests currently open for category {0}, shipper {1}", category, shipperSymbol);
				return;
			}

			foreach (ITDCCloseOutItem manifest in list)
			{
				// Close out
				Console.WriteLine("\nClosing out manifest: {0} ({1})", manifest.Symbol, manifest.FriendlyName);
				ret = transAPI.CloseOut(category, shipperSymbol, manifest);
				if (ret.Code != (int)ErrorCode.ecNoError)
				{
					Console.WriteLine("CloseOut returned: {0} ({1}).", ret.Message, ret.Code);
					return;
				}

				// Get the CloseOut results
				ITDCCloseOutResult closeoutresult = (ITDCCloseOutResult)ret.Data;

				// Display the manifest totals
				Console.WriteLine("Manifest totals:");
				ITransAPIData resultdata = closeoutresult.ResultData;
				displayTransAPIData(resultdata);

				// Display the ShipFileItem data
				ITDCShipFileItem shipfileitem = closeoutresult.ShipFileItem;
				Console.WriteLine("ShipFileItem data:");
				Console.WriteLine("     ITDCShipFileItem.Symbol");
				Console.WriteLine("         {0}", shipfileitem.Symbol);
				Console.WriteLine("     ITDCShipFileItem.FriendlyName");
				Console.WriteLine("         {0}", shipfileitem.FriendlyName);
				Console.WriteLine("     ITDCShipFileItem.SequenceNumber");
				Console.WriteLine("         {0}", shipfileitem.SequenceNumber);

				// Display the ShipFile Attributes
				Console.WriteLine("ShipFileItem attributes:");
				ITransAPIData shipfileattributes = shipfileitem.Attributes;
				displayTransAPIData(shipfileattributes);

				// Display TransmitItem data
				ITDCCollection transmititemlist = closeoutresult.TransmitItemList;
				int x = 0;
				foreach (ITDCTransmitItem item in transmititemlist)
				{
					x++;
					Console.WriteLine("Transmit Item #{0} of {1}:", x, transmititemlist.Count);
					Console.WriteLine("     ITDCTransmitItem.Symbol");
					Console.WriteLine("         {0}", item.Symbol);
					Console.WriteLine("     ITDCTransmitItem.FriendlyName");
					Console.WriteLine("         {0}", item.FriendlyName);
					Console.WriteLine("     ITDCTransmitItem.SequenceNumber");
					Console.WriteLine("         {0}", item.SequenceNumber);
					Console.WriteLine("     ITDCTransmitItem.TransmitStatus");
					Console.WriteLine("         {0}", item.TransmitStatus);
				}

				// print end-of-day document
				if (generateDocuments)
				{
					DocumentHandler docHandler = new DocumentHandler(transAPI);
					string formatSymbol = "CONNECTSHIP_UPS_PICKUP_SUMMARY_BARCODE.STANDARD";
					docHandler.GenerateAndPrintDocument(LogicalDocumentType.ldtShipFile, formatSymbol, shipfileitem.SequenceNumber, category, shipperSymbol, shipfileitem);
				}
			}
		}

		private static void updateDB(string orderId, ITransAPIPackageResult packageResult)
		{
			using (SqlConnection dbConn = new SqlConnection(sqlConnectionString))
			{
				dbConn.Open();

				ITDCMoney total = (ITDCMoney)packageResult.ResultData[F.APPORTIONED_TOTAL];
				ITDCIdentity shipper = (ITDCIdentity)packageResult.ResultData[F.SHIPPER];
				int msn = (int)packageResult.ResultData[F.MSN];

				string trackingNumber = Convert.ToString(packageResult.ResultData[F.TRACKING_NUMBER]);
				if (trackingNumber.Length == 0)
					trackingNumber = Convert.ToString(packageResult.ResultData[F.BAR_CODE]);
				if (trackingNumber.Length == 0)
					trackingNumber = Convert.ToString(packageResult.ResultData[F.WAYBILL_BOL_NUMBER]);
				if (trackingNumber.Length == 0)
					trackingNumber = "0";

				using (SqlCommand cmd = new SqlCommand("UPDATE [orders] SET [db_tracking] = @trackingNumber, [db_total] = @total, [db_msn] = @msn, [db_shipper] = @shipper, [db_shipdate] = @shipdate WHERE [db_order] = @orderId", dbConn))
				{
					cmd.Parameters.AddWithValue("@trackingNumber", trackingNumber);
					cmd.Parameters.AddWithValue("@total", total.Amount);
					cmd.Parameters.AddWithValue("@msn", msn);
					cmd.Parameters.AddWithValue("@shipper", shipper.Symbol);
					cmd.Parameters.AddWithValue("@shipdate", ShipPrint.shipDate.FormatString);
					cmd.Parameters.AddWithValue("@orderId", orderId);
					int rowsAffected = cmd.ExecuteNonQuery();
					if (rowsAffected == 1)
					{
						Console.WriteLine("Order id {0} updated: db_tracking = {1}, db_total = {2}, db_msn = {3}, db_shipper = {4}, db_shipdate = {5}", orderId, trackingNumber, total.FormatStringEx, msn, shipper.Symbol, ShipPrint.shipDate.FormatStringEx);
					}
					else if (rowsAffected == 0)
					{
						Console.Error.WriteLine("*** Unable to find order id {0} ***", orderId);
					}
				}
			}
		}

		private static string getProgisticsService(string db_service)
		{
			switch (db_service)
			{
				case "UPSG": return "CONNECTSHIP_UPS.UPS.GND";
				case "UPS3": return "CONNECTSHIP_UPS.UPS.3DA";
				case "UPS2A": return "CONNECTSHIP_UPS.UPS.2AM";
				case "UPS2": return "CONNECTSHIP_UPS.UPS.2DA";
				case "UPS1": return "CONNECTSHIP_UPS.UPS.NDA";
				case "UPS1A": return "CONNECTSHIP_UPS.UPS.NAM";
				case "UPSCA": return "CONNECTSHIP_UPS.UPS.STD";
				case "UPSW3": return "CONNECTSHIP_UPS.UPS.EPD";
				case "UPSW1": return "CONNECTSHIP_UPS.UPS.EXP";
				case "UPSW2": return "CONNECTSHIP_UPS.UPS.EXPSVR";
				case "ONGHT": return "CONNECTSHIP_GLOBAL.SPEEDEE.ONGHT";
				default:
					return null;
			}
		}

		// Map order database country to Progistics country symbol
		private static string getCountrySymbol(string db_country)
		{
			switch (db_country)
			{
				case "USA": return "UNITED_STATES";
				case "Canada": return "CANADA";
				case "France": return "FRANCE";
				case "Belgium": return "BELGIUM";
				case "Germany": return "GERMANY";
				default:
					return null;
			}
		}

		private static void addInternationalAttributes(CoDictionary package)
		{
			var contents = new CoSimpleList();
			for (int i = 1; i <= 1; i++)
			{
				var content = new CoDictionary();
				content[F.QUANTITY] = 1;
				content[F.QUANTITY_UNIT_MEASURE] = "PCS";
				content[F.UNIT_VALUE] = 50;
				content[F.UNIT_WEIGHT] = .1;
				content[F.LICENSE_NUMBER] = "NLR";
				content[F.DESCRIPTION] = "Bar-B-Que Sauce " + i;
				content[F.EXPORT_HARMONIZED_CODE] = "2103.90.9091";
				content[F.EXPORT_INFORMATION_CODE] = "DD";
				content[F.ORIGIN_COUNTRY] = "UNITED_STATES";
				content[F.PRODUCT_CODE] = "54321";
				contents.Add(content);
			}
			package[F.COMMODITY_CONTENTS] = contents;

			// Generate psuedo-AES trans no. ("AES X<year><month><day>123456")
			package[F.AES_TRANSACTION_NUMBER] = DateTime.Now.ToString("AES XyyyyMMdd123456");

			package[F.SED_METHOD] = SEDMethod.sedmElectronicallyFiled;
		}

		private static void generateAndPrintDocument(ITransAPI transAPI, int msn, int bundleId, string serviceSymbol, string shipper)
		{
			ITDCReturn ret;

			// Print documents for packages from MSNs we saved during shipping
			Console.WriteLine("\nPrinting Labels");
			DocumentHandler docHandler = new DocumentHandler(transAPI);

			// Use the "SERVER.CATEGORY" from the "SERVER.CATEGORY.SUBCATEGORY" service symbol
			string category = serviceSymbol.Substring(0, serviceSymbol.LastIndexOf('.'));

			// Find the package data for this msn, check to see if the consignee is Wal-Mart
			ISimpleList searchResultFields = new CoSimpleList();
			searchResultFields.Add(F.CONSIGNEE);
			CoDictionary filter = new CoDictionary();
			filter[F.MSN] = msn;

			ret = transAPI.SearchPackages(category, filter, searchResultFields, SearchCloseOutMode.scomRelease);
			if (ret.IsSuccess)
			{
				ITDCSearchPackagesResult result = (ITDCSearchPackagesResult)ret.Data;
				while (result.MoreItems)
				{
					ITDCSearchPackagesItem item = result.NextItem;
					CoNameAddress consignee = (CoNameAddress)item.ResultData[F.CONSIGNEE];
					if (consignee.Company.ToLower().Contains("wal-mart"))
					{
						string walmartFormatSymbol = "WALMART_PACKAGE.STANDARD";
						docHandler.GenerateAndPrintDocument(LogicalDocumentType.ldtPackage, walmartFormatSymbol, msn, category, shipper);
					}
					Console.WriteLine();
					break;
				}
			}

			// generate Global SAMPLE_SPEEDEE.STANDARD custom label
			if (serviceSymbol.Equals("CONNECTSHIP_GLOBAL.SPEEDEE.ONGHT"))
			{
				docHandler.GenerateAndPrintDocument(LogicalDocumentType.ldtPackage, "SAMPLE_SPEEDEE.STANDARD", msn, category, shipper);
				return;
			}

			// Find first package document and format that can be generated for this category/carrier
			Console.WriteLine("Finding first document.format that can be generated...");

			// Get all the documents for this category
			ret = transAPI.GetDocuments(category);

			bool docGenerated = false;

			foreach (ITDCIdentity document in (ITDCIdentityCollection)ret.Data)
			{
				Console.WriteLine(" Document: {0} ({1})", document.FriendlyName, document.Symbol);

				// Get all the formats for this document
				ret = transAPI.GetDocumentFormats(category, document.Symbol);
				foreach (ITDCIdentity format in (ITDCIdentityCollection)ret.Data)
				{
					Console.WriteLine("    Format: {0} ({1})", format.FriendlyName, format.Symbol);

					// Attempt to generate a document for the package
					LogicalDocumentType ldt = transAPI.GetLogicalDocumentType(category, document.Symbol);
					int documentID = ldt == LogicalDocumentType.ldtBundle ? bundleId : msn;

					ret = docHandler.GenerateAndPrintDocument(ldt, format.Symbol, documentID, category, shipper);

					if (ret.IsSuccess)
						docGenerated = true;
					else
						break; // if unable to generate document, don't try to generate another document of this type for the next format

					if (docGenerated)
						break;
				}
				if (docGenerated)
					break;
			}
		}

		// Helper methods to display result data

		// Display rating or shipping results for a single service
		private static void displayTransAPIData(ITransAPIData transapiData, int indent = 4)
		{
			string indentation = new String(' ', indent);
			Console.WriteLine(indentation + "{0} Count: ({1})", transapiData.GetType().Name, transapiData.Count);
			foreach (string key in transapiData)
				displayKeyValue(key, transapiData[key], indent);
		}

		// Display dictionary contents
		private static void displayDictionaryData(ITDCDictionary dictionary, int indent = 4)
		{
			string indentation = new String(' ', indent);
			Console.WriteLine(indentation + "{0} Count: ({1})", dictionary.GetType().Name, dictionary.Count);
			foreach (string key in dictionary)
				displayKeyValue(key, dictionary[key], indent);
		}

		// Display a single key/value pair from a dictionary
		private static void displayKeyValue(string key, object data, int indent)
		{
			string indentation = new String(' ', indent);

			// Display the name and type of the item
			try
			{
				Console.ForegroundColor = ConsoleColor.Cyan;
				Console.WriteLine(indentation + "{0} ({1})", key, data.GetType().Name);
			}
			finally
			{
				Console.ResetColor();
			}
			displayValue(data, indent + 4);
		}

		private static void displayValue(object data, int indent)
		{
			try
			{
				Console.ForegroundColor = ConsoleColor.Yellow;

				string indentation = new String(' ', indent);

				if (data is ISimpleList)
				{
					ISimpleList list = data as ISimpleList;
					//Console.WriteLine(indentation + "Count: ({0})", list.Count);
					foreach (object o in list)
						displayValue(o, indent + 2);
				}
				else if (data is ITDCDictionary)
				{
					displayDictionaryData((ITDCDictionary)data, indent + 2);
				}
				else if (data is ITDCCommitment)
				{
					ITDCCommitment committment = data as ITDCCommitment;
					Console.WriteLine(indentation + "{0} ({1})", committment.FriendlyName, committment.Symbol);
					Console.WriteLine(indentation + "Committment days: {0}", committment.Days);
					Console.WriteLine(indentation + "Committment time: {0}", committment.Time);
				}
				else if (data is ITDCDate)
				{
					ITDCDate date = data as ITDCDate;
					Console.WriteLine(indentation + "{0}", date.FormatStringEx);
				}
				else if (data is ITDCWeight)
				{
					ITDCWeight weight = data as ITDCWeight;
					Console.WriteLine(indentation + "{0}", weight.FormatStringEx);
				}
				else if (data is ITDCMoney)
				{
					ITDCMoney money = data as ITDCMoney;
					Console.WriteLine(indentation + "{0}", money.FormatStringEx);
				}
				else if (data is ITDCDimension)
				{
					ITDCDimension dimension = data as ITDCDimension;
					Console.WriteLine(indentation + "{0}", dimension.FormatStringEx);
				}
				else if (data is ITDCIdentity)
				{
					ITDCIdentity identity = data as ITDCIdentity;
					Console.WriteLine(indentation + "Symbol: {0}, FriendlyName: {1}", identity.Symbol, identity.FriendlyName);
				}
				else if (data is ITDCNameAddress)
				{
					ITDCNameAddress na = data as ITDCNameAddress;
					Console.WriteLine(indentation + "Company:       {0}", na.Company);
					Console.WriteLine(indentation + "Address1:      {0}", na.Address1);
					Console.WriteLine(indentation + "Address2:      {0}", na.Address2);
					Console.WriteLine(indentation + "City:          {0}", na.City);
					Console.WriteLine(indentation + "StateProvince: {0}", na.StateProvince);
					Console.WriteLine(indentation + "CountrySymbol: {0}", na.CountrySymbol);
					Console.WriteLine(indentation + "PostalCode:    {0}", na.PostalCode);
					Console.WriteLine(indentation + "Residential:   {0}", na.Residential);
				}
				else
				{
					Console.WriteLine(indentation + "{0}", data);
				}

			}
			finally
			{
				Console.ResetColor();
			}
		}
	}

	// Helper to print document

	public enum DocumentOutput
	{
		Pdf,
		Png,
		Device,
	}


	public class DocumentHandler
	{
		// Physical Printer Options
		private const string port = @"lpd://trainingsql/trngeltron";
		private const string deviceSymbol = "ELTRON.ELTRON2442";

		// Stock
		private const string stockSymbol = "THERMAL_LABEL_8"; // 4" x 6" Thermal Label - height 6000, width 4000

		// instance variables
		private ITransAPI transAPI = null;
		private ITDCPrinter physicalPrinter = null;
		private CoDeviceManagerDirect physicalDeviceManager = null;

		public ITDCDocumentDestination DocumentDestination { get; private set; }

		public DocumentHandler(ITransAPI transAPI)
		{
			this.transAPI = transAPI;
			this.initialize();
		}


		private void initialize()
		{
			ITDCStockDescriptor stockdescriptor = (ITDCStockDescriptor)transAPI.PrinterStocks[stockSymbol];

			this.initializePhysicalPrinterIfAvailable(stockdescriptor);

			if (this.DocumentDestination == null)
			{
				// Physical printer not available; set up pseudo-device and indicate what physical document types are supported
				this.DocumentDestination = new CoDocumentDestination();
				this.DocumentDestination.Stock = stockdescriptor;
				this.DocumentDestination.SupportedDocumentTypes = (int)(DocumentType.dtPosition | DocumentType.dtBinary);
			}
		}


		[Conditional("DEVICE_AVAILABLE")]
		private void initializePhysicalPrinterIfAvailable(ITDCStockDescriptor stockdescriptor)
		{
			// Initialize Device Manager
			this.physicalDeviceManager = new CoDeviceManagerDirect();
			ITDCReturn ret = this.physicalDeviceManager.OpenDeviceDirect(port, deviceSymbol, null);
			if (!ret.IsSuccess)
				throw new ApplicationException(String.Format("Printer Initialization Error: ({0}) {1}", ret.Code, ret.Message));

			// Obtain the initialized printer object
			this.physicalPrinter = (ITDCPrinter)ret.Data;

			ret = this.physicalPrinter.SetStock(stockdescriptor);
			if (!ret.IsSuccess)
				throw new ApplicationException(String.Format("Printer SetStock Error: ({0}) {1}", ret.Message, ret.Code));

			// Set the document destination
			ret = this.physicalPrinter.DocumentDestination;
			if (!ret.IsSuccess)
				throw new ApplicationException(String.Format("Printer DocumentDestination Error: ({0}) {1}", ret.Message, ret.Code));

			// Get the Document Destination object
			this.DocumentDestination = (ITDCDocumentDestination)ret.Data;
		}


		public ITDCReturn GenerateAndPrintDocument(LogicalDocumentType ldt, string formatSymbol, int documentID, string category, string shipper, ITDCShipFileItem shipFileItem = null)
		{
			ITDCReturn ret = new CoReturn();
			IDictionary docParameters = new CoDictionary();

			string identifierName;
			switch (ldt)
			{
				case LogicalDocumentType.ldtPackage:
					identifierName = "MSN";
					break;
				case LogicalDocumentType.ldtBundle:
					identifierName = "BUNDLE";
					break;
				case LogicalDocumentType.ldtShipFile:
					if (shipFileItem == null)
					{
						string error = "ShipFileItem required for document " + formatSymbol;
						Console.WriteLine(error);
						return new CoReturn(ErrorCode.ecTransAPINoDataForDocument, error);
					}
					identifierName = "SHIPFILE";
					break;

				default:
					return new CoReturn(ErrorCode.ecUnknown, String.Format("Unexpected logical document type - {0}", ldt));
			}

			Console.WriteLine("Generating document {0} for {1} {2}", formatSymbol, identifierName, documentID);
			switch (ldt)
			{
				case LogicalDocumentType.ldtPackage:
					ret = transAPI.GeneratePackageDocument(category, shipper, formatSymbol, this.DocumentDestination, documentID, docParameters);
					break;
				case LogicalDocumentType.ldtBundle:
					ret = transAPI.GenerateBundleDocument(category, shipper, formatSymbol, this.DocumentDestination, documentID, docParameters);
					break;
				case LogicalDocumentType.ldtShipFile:
					ret = transAPI.GenerateShipFileDocument(category, shipper, formatSymbol, this.DocumentDestination, shipFileItem, docParameters);
					break;
			}

			if (ret.IsSuccess)
			{
				Console.WriteLine("Printing document {0} for {1} {2}", formatSymbol, identifierName, documentID);
				this.printDocument((ITDCDocument)ret.Data, identifierName, documentID);
			}
			else
				Console.WriteLine("Unable to generate document {0} - {1} ({2})", formatSymbol, ret.Message, ret.Code);

			return ret;
		}


		private void printDocument(ITDCDocument generatedDocument, string identifierName, int documentID)
		{
			DocumentOutput docOutput = ShipPrint.documentOutput;
			this.setDocumentOutput(ref docOutput);
			switch (docOutput)
			{
				case DocumentOutput.Device:
					this.printPositionDocument(generatedDocument, displayBinary: false);
					break;
				case DocumentOutput.Png:
					this.savePositionDocumentToPngFile(generatedDocument, identifierName, documentID);
					break;
				case DocumentOutput.Pdf:
					this.savePositionDocumentToPdfFile(generatedDocument, identifierName, documentID);
					break;
				default:
					break;
			}
		}


		[Conditional("DEVICE_AVAILABLE")]
		private void setDocumentOutput(ref DocumentOutput docOutput)
		{
			docOutput = DocumentOutput.Device;
		}


		[Conditional("DEVICE_AVAILABLE")]
		private void printPositionDocument(ITDCDocument generatedDocument, bool displayBinary)
		{
			// Start the document
			ITDCReturn ret = this.physicalPrinter.StartDocument("My Document");
			if (!ret.IsSuccess)
			{
				Console.WriteLine("StartDocument Error: ({0}) {1}", ret.Message, ret.Code);
				return;
			}

			if (positionDocAvailable(generatedDocument))
			{
				ITDCPositionDocument positiondocument = generatedDocument.PositionDocument;
				foreach (ITDCPositionDocumentItem item in positiondocument.ItemList)
				{
					if (displayBinary)
					{
						// Convert document to binary and display in command prompt
						ret = this.physicalPrinter.ConvertPositionToBinary(item);
						if (ret.IsSuccess)
							Console.WriteLine("Binary version of document:\n" + ret.Data);
					}

					Console.WriteLine("Sending document to {0}", this.physicalPrinter.FriendlyName);

					// Send to printer
					ret = this.physicalPrinter.PrintPosition(item);
					if (!ret.IsSuccess)
					{
						Console.WriteLine("PrintPosition Error: ({0}) {1}", ret.Message, ret.Code);
						return;
					}
				}
			}

			// Complete document printing
			ret = this.physicalPrinter.EndDocument();
			if (!ret.IsSuccess)
			{
				Console.WriteLine("EndDocument Error: ({0}) {1}", ret.Message, ret.Code);
				return;
			}
		}


		private void savePositionDocumentToPngFile(ITDCDocument generatedDocument, string identifierName, int documentID)
		{
			if (positionDocAvailable(generatedDocument))
			{
				ITDCPositionDocument positiondocument = generatedDocument.PositionDocument;
				foreach (ITDCPositionDocumentItem item in positiondocument.ItemList)
				{
					ITDCReturn ret = item.GetImageWithDD(this.DocumentDestination, true);
					if (!ret.IsSuccess)
					{
						Console.WriteLine("Error getting image: ({0}) {1}", ret.Message, ret.Code);
						return;
					}

					// Get image and write to local image file
					ITDCImage image = (ITDCImage)ret.Data;

					string imageFile = this.getFileName(identifierName, documentID, ".png");
					Console.WriteLine("Saving {0}", imageFile);

					ret = image.SaveAsFile(imageFile, ImageType.itPNG);
					if (!ret.IsSuccess)
						Console.WriteLine("Error saving image: ({0}) {1}", ret.Message, ret.Code);
					else // Launch the default viewer
						Process.Start(imageFile);
				}
			}
		}


		private void savePositionDocumentToPdfFile(ITDCDocument generatedDocument, string identifierName, int documentID)
		{
			if (positionDocAvailable(generatedDocument))
			{
				ITDCPositionDocument positiondocument = generatedDocument.PositionDocument;
				ITDCReturn ret = positiondocument.GetPDF(this.DocumentDestination, false);
				if (!ret.IsSuccess)
				{
					Console.WriteLine("Error getting PDF: ({0}) {1}", ret.Message, ret.Code);
					return;
				}

				// write to local PDF file
				ITDCPDFDocument pdf = (ITDCPDFDocument)ret.Data;

				string pdfFile = this.getFileName(identifierName, documentID, ".pdf");
				Console.WriteLine("Saving {0}", pdfFile);

				ret = pdf.SaveAsFile(pdfFile);
				if (!ret.IsSuccess)
					Console.WriteLine("Error saving PDF: ({0}) {1}", ret.Message, ret.Code);
				else // Launch the default viewer
					Process.Start(pdfFile);
			}
		}


		private bool positionDocAvailable(ITDCDocument doc)
		{
			return (doc.AvailableDocumentTypes & (int)DocumentType.dtPosition) != 0;
		}


		private string getFileName(string identifierName, int documentID, string extension)
		{
			string tempFile = Path.GetTempFileName();
			return Path.Combine(Path.GetDirectoryName(tempFile), Path.GetFileNameWithoutExtension(tempFile) + String.Format("_{0}_{1}", identifierName, documentID)) + extension;
		}


		~DocumentHandler() // destructor
		{
			if (this.physicalDeviceManager != null)
			{
				this.physicalDeviceManager.CloseDeviceDirect(this.physicalPrinter);
				this.physicalDeviceManager.DestroyNotify();
			}
		}
	}
}