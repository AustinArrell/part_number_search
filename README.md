# part_number_search

This is a simple tool that, when given a list of Copy Machine part numbers, it will give you a list of model numbers that the part is compatible with. 
I created this while I was a repair technician for Ricoh and I was helping my manager do inventory on a couple boxes of loose, unknown parts where the only available data was the part number

It functions by making web requests to a site named Precision Roller and scrapes the return HTML for the model numbers

If I did this again today or if I revisit someday I would look for an api or perhaps cache the scraped data into a database so that the app is not 100% relying on a 3rd party website for the live data.
