# emailtocsvconverter
A simple class converting an e-mail directory or a single e-mail file to a CSV formatted output

Usage (PHP Version):

1. Download the files and place them into your webserver directory.
2. Launch using: http://localhost/converter.html
3. If needed modify the path used in converter.html.

If You run into out of memory problems I recommend setting memory_limit in php.ini to 16MB. 
You also may need to modify max_execution_time parameter to 360 seconds or longer in php.ini to be able to process a large dataset. 
In my case I was running a 1.3GB+ dataset and the script took 5.6300687829653 minutes complete.

Usage (Python Version):

1. Download converter.py
2. Launch using: converter.py --input_filename <your source filename or directory example: input.txt or INPUT> --output_filename <your output filename example: output.csv>