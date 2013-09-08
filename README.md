Email XSLX to Convert to CSV, JSON and XML
==========================

This is a working project that allows the sending of Microsoft Excel files to an email address and automatically having them converted to CSV, JSON and XML.

I wanted a dead simple way to manage conversion of the flood of spreadsheets I was finding in government, but also handle it in a way that can be taught to any potential data steward.

This is a ongoing project, that I will work on when I have time and the need. I encourage your participation.

## Dependencies:

PHP 5+
PHPExcel - http://phpexcel.codeplex.com/
Gmail
IMAP
PHPMailer - http://phpmailer.worxware.com/

## Details

This is a single, rough script that will check a Gmail email address using IMAP (can be adjusted to handle any POP as well), downloads the attachment. Then using PHPExcel it converts to CSV and secondarily to JSON then XML.  

It uses a JSON file to decide who has access to be sending emails to convert, as well as a JSON manifest of files it has process.

It records each attachment in manifest before it processes, because of the size of some files and other potential errors it still requires intervention, and scaling of the Amazon EC2 instance I run this on at night to handle the memory load.

Upon successful processing it sends an email back to sender with links to CSV, JSON and XML formats.

## Roadmap

* Autoscaling of higher load to larger EC2 instance
* upload web interface
* other email possibilities beyond gmail / imap
* ec2 appliance (AMI)

## Contact

Kin Lane

kinlane@gmail.com

@kinlane
