# Release to Commerce Report
An Excel VBA project that creates a quick overview of all the products manufactured at the Pittsburgh site which are completed and ready to be released and shipped. Shows detailed information such as dollar value of the product, related status, and location in the warehouse. Data is retrieved from JD Edwards EnterpriseOne.

This report was created to replace the existing one that had the following problems:
* Slow
* Had missing information about products
* Total value of a product was not always accurate
This project is done in Excel since it was the only tool available at work computers to complete such a report. Similar report can be done inside JD Edwards EnterpriseOne, though without any fancy graphics. Also there are better tools for this task, for example Power BI Desktop, Tableau, etc.

## Usage
See the project [wiki](https://github.com/ykoziy/release-to-commerce-report/wiki) for detailed usage instructions.

## Limitations
  * Slow, while it is much faster when compared to the original it is still slow. Since data has to be retrieved from multiple columns in multiple tables across the database and then dumped into excel file. If data grows, this would be a major issue.
  * Not fully automatic, some user interaction is required. End goal would be completely automating the process so that employees would get a fresh report every morning via their email.

## License
[MIT](https://github.com/ykoziy/release-to-commerce-report/blob/master/LICENSE)
