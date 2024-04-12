# Nav To Nop Stock Sync App

The NopCommerce-DynamicsNav-StockSync project is a console application built on the .NET Framework 4.6.1. It facilitates seamless integration between NopCommerce, an open-source e-commerce platform, and Microsoft Dynamics Nav ERP using OData services.

### Features:
- **ERP Integration:** Connects NopCommerce with Microsoft Dynamics Nav ERP to synchronize stock data.
- **Batch Processing:** Fetches stock data in batches to accommodate the service's limit of 1000 records per request.
- **Error Handling:** Implements retry logic for network errors, ensuring uninterrupted data synchronization.
- **Data Extraction:** Extracts product SKU, location code, and stock quantity from Dynamics Nav ERP.
- **Database Interaction:** Prepares data tables to update NopCommerce database with accurate stock quantities.
- **Reporting:** Generates an Excel report summarizing all product stock updates for easy reference.
- **Email Notification:** Can send the generated report to the relevant department via email for further analysis.

### Usage:
1. Clone the repository to your local machine.
2. Configure the application settings with the necessary credentials and endpoints for NopCommerce and Dynamics Nav ERP.
3. Run the console application.
4. Use Debugger to monitor the synchronization process as it fetches, processes, and updates stock data.
5. Review the generated Excel report for insights into product stock updates.

Streamline your stock synchronization process between NopCommerce and Dynamics Nav ERP with NopCommerce-DynamicsNav-StockSync. Ensure accurate stock management and optimize your e-commerce operations effortlessly.
