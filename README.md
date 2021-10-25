# nmap_xml_parser

[Nmap](https://nmap.org/)(Network Mapper) can provide different types of outputs, and one of them is XML formatted output. 

This tool processes XML formatted output of Nmap by parsing IP addresses, open ports and the services run on these ports. Then the result is written in a table in Microsoft Word (.docx) document. Thus, some valuable information is parsed from the Nmap output and stored in a Word document (.docx) in a beutified format.

## Usage

`python nmap_xml_parser.py -f <nmap_xml_output_file> -o <output_file>`

**Example:** <br>
`python nmap_xml_parser.py -f nmap_result.xml -o parsed_result`

As a result of the command above, a file named ***parsed_result.docx*** is created.

**How Does The Output Look Like?**<br>
![](https://raw.githubusercontent.com/yunemse48/nmap_xml_parser/main/img/output_ss.png)
