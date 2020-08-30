# Sharepoint crossplatform multithreaded file management tool
  
[![License](http://img.shields.io/badge/license-mit-blue.svg?style=flat-square)](https://raw.githubusercontent.com/json-iterator/go/master/LICENSE)
[![Build Status](https://travis-ci.org/gvaduha/sharepoint-filemgmt-tool.svg?branch=master)](https://travis-ci.org/gvaduha/sharepoint-filemgmt-tool)

## Brief
Cross platform multithreaded command line utility to manage files on Sharepoint (tm) server.

## Purpose
Sometimes you have to share files, especially large using Sharepoint.
It's mostly diabolic practic but we have to obey the rules.
File management in Sharepoint typically has inhuman interface but system administrators tends to restrict it more to have misirable file size limit and force to upload files one by one.
This is untolerable for most human beings and totaly outrageous act for software developers.

It has started as minimal "what-i-need-only" tool. Code now needs refactoring to split up into smaller function blocks. More ls options, progress report features (wget like) along with extended functions may be included in future versions or done by request.

## Examples
upload:
```sh
shrpnt-fm -s https://xxx.xxx/sites/MySite -f "Shared Documents/MyFolder" -u Iam -p Mypass --up file*
```
download:
```sh
shrpnt-fm -s https://xxx.xxx/sites/MySite -f "Shared Documents/MyFolder" -u Iam -p Mypass --down file*
```
remove:
```sh
shrpnt-fm -s https://xxx.xxx/sites/MySite -f "Shared Documents/MyFolder" -u Iam -p Mypass --rm tracing.?.log
```
list:
```sh
shrpnt-fm -s https://xxx.xxx/sites/MySite -f "Shared Documents/MyFolder" -u Iam -p Mypass
shrpnt-fm -s https://xxx.xxx/sites/MySite -f "Shared Documents/MyFolder" -u Iam -p Mypass --ls Docum*
```
