# powerHermod
powerHermod (https://en.wikipedia.org/wiki/Herm%C3%B3%C3%B0r) is a tool to create powerpoint presentation with images from wms services and websites.

Specific used to generate weather forecast briefings from the website:
http://ifm.fcoo.dk/

Using the commandline tool cutycapt to take a "screenshot" of a website.

Can also download a specific image from a website, it the URL can be calculated or
is static.

## Using
THe package consist of Grane.py (https://en.wikipedia.org/wiki/Grani), and several configuration files.

grane.yaml is the main configuration file.
A folder with several configuration file, one for each presentation.

The best way to use powerHermod is to have a cron jobs, which runs each hour.

# 2017-02-21
Added support for multi configuration files in folder specified in Grane.yaml.


# Roadmap ahead
(2017-02-18)
Extended to generate pdf files.
Extended to attach the images as attachment to the presentation email.
Extended to use the wms service direct.

