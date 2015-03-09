#!/usr/bin/python

# This script was inspired by:
# http://blogs.msdn.com/b/exchangedev/archive/2009/02/05/quick-and-dirty-unix-shell-scripting-with-ews.aspx
# http://ewsmacwidget.codeplex.com/

import os
from lxml import etree
from datetime import datetime
from datetime import date
from datetime import timedelta
from pytz import timezone
from StringIO import StringIO
import pytz
import pycurl
import base64
import ConfigParser

# Read the config file
timezoneLocation = os.getenv('TZ', 'UTC')
config = ConfigParser.RawConfigParser({
  'path': '/ews/Exchange.asmx',
  'username': '',
  'password': '',
  'auth_type': 'any',
  'cainfo': '',
  'timezone': timezoneLocation,
  'days_history': 7,
  'days_future': 30,
  'max_entries': 100})
dir = os.path.dirname(os.path.realpath(__file__))
config.read(os.path.join(dir, 'config.cfg'))

# Exchange user and password
ewsHost = config.get('ews-orgmode', 'host')
ewsUrl = config.get('ews-orgmode', 'path')
ewsUser = config.get('ews-orgmode', 'username')
ewsPassword = config.get('ews-orgmode', 'password')
ewsAuthType = config.get('ews-orgmode', 'auth_type').lower()
ewsCAInfo = config.get('ews-orgmode', 'cainfo')
timezoneLocation = config.get('ews-orgmode', 'timezone')
daysHistory = config.getint('ews-orgmode', 'days_history')
daysFuture = config.getint('ews-orgmode', 'days_future')
maxEntries = config.getint('ews-orgmode', 'max_entries')

def parse_ews_date(dateStr):
  d = datetime.strptime(dateStr, "%Y-%m-%dT%H:%M:%SZ")
  exchangeTz = pytz.utc
  localTz = timezone(timezoneLocation)
  return exchangeTz.localize(d).astimezone(localTz);

def format_orgmode_date(dateObj):
  return dateObj.strftime("%Y-%m-%d %H:%M")

def format_orgmode_time(dateObj):
  return dateObj.strftime("%H:%M")

# Helper function to write an orgmode entry
def print_orgmode_entry(subject, start, end, location, response):
  startDate = parse_ews_date(start);
  endDate = parse_ews_date(end);
  # Check if the appointment starts and ends on the same day and use proper formatting
  dateStr = ""
  if startDate.date() == endDate.date():
    dateStr = "<" +  format_orgmode_date(startDate) + "-" + format_orgmode_time(endDate) + ">"
  else:
    dateStr = "<" +  format_orgmode_date(startDate) + ">--<" + format_orgmode_date(endDate) + ">"

  if subject is not None:
    if dateStr != "":
      print "* " + dateStr + " " + subject.encode('ascii', 'ignore')
    else:
      print "* " + subject.encode('ascii', 'ignore')

  if location is not None:
    print ":PROPERTIES:"
    print ":LOCATION: " + location.encode('utf-8')
    print ":RESPONSE: " + response.encode('utf-8')
    print ":END:"

  print ""

# Debug code
# print_orgmode_entry("subject", "2012-07-27T11:10:53Z", "2012-07-27T11:15:53Z", "location", "participants")
# exit(0)

# Build the soap request
# For CalendarItem documentation, http://msdn.microsoft.com/en-us/library/exchange/aa564765(v=exchg.140).aspx
start = date.today() - timedelta(days=daysHistory)
end = date.today() + timedelta(days=daysFuture)
request = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xmlns:xsd="http://www.w3.org/2001/XMLSchema"
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindItem Traversal="Shallow" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <ItemShape>
        <t:BaseShape>Default</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:MyResponseType"/>
        </t:AdditionalProperties>
      </ItemShape>
      <CalendarView MaxEntriesReturned="{2}" StartDate="{0}T00:00:00-08:00" EndDate="{1}T00:00:00-08:00"/>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </ParentFolderIds>
    </FindItem>
  </soap:Body>
</soap:Envelope>""".format(start, end, maxEntries)
request_len = len(request)
request = StringIO(request)
# Debug code
# print request.getvalue()
# exit(0)

h = []
h.append('Content-Type: text/xml; charset=UTF-8')
h.append('Content-Length: %d' % request_len);

header = StringIO()
body = StringIO()

c = pycurl.Curl()
# Debug code
# c.setopt(c.VERBOSE, 1)

c.setopt(c.URL, 'https://%s%s' % (ewsHost, ewsUrl))
c.setopt(c.POST, 1)
c.setopt(c.HTTPHEADER, h)

if ewsAuthType == 'digest':
  c.setopt(c.HTTPAUTH, c.HTTPAUTH_DIGEST)
elif ewsAuthType == 'basic':
  c.setopt(c.HTTPAUTH, c.HTTPAUTH_BASIC)
elif ewsAuthType == 'ntlm':
  c.setopt(c.HTTPAUTH, c.HTTPAUTH_NTLM)
elif ewsAuthType == 'negotiate':
  c.setopt(c.HTTPAUTH, c.HTTPAUTH_GSSNEGOTIATE)
elif ewsAuthType == 'any':
  c.setopt(c.HTTPAUTH, c.HTTPAUTH_ANYSAFE)
c.setopt(c.USERPWD, '%s:%s' % (ewsUser, ewsPassword))

if len(ewsCAInfo) > 0:
  c.setopt(c.CAINFO, ewsCAInfo)

# http://stackoverflow.com/questions/27808835/fail-to-assign-io-object-to-writedata-pycurl
c.setopt(c.WRITEFUNCTION, body.write)
c.setopt(c.HEADERFUNCTION, header.write)
c.setopt(c.READFUNCTION, request.read)

c.perform()
c.close()

# Read the webservice response
data = body.getvalue()

# Debug code
# print data
# exit(0)

# Parse the result xml
root = etree.fromstring(data)

xpathStr = "/s:Envelope/s:Body/m:FindItemResponse/m:ResponseMessages/m:FindItemResponseMessage/m:RootFolder/t:Items/t:CalendarItem"
namespaces = {
    's': 'http://schemas.xmlsoap.org/soap/envelope/',
    't': 'http://schemas.microsoft.com/exchange/services/2006/types',
    'm': 'http://schemas.microsoft.com/exchange/services/2006/messages',
}

# Print calendar elements
elements = root.xpath(xpathStr, namespaces=namespaces)
for element in elements:
  subject= element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Subject').text
  location= element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Location').text
  start = element.find('{http://schemas.microsoft.com/exchange/services/2006/types}Start').text
  end = element.find('{http://schemas.microsoft.com/exchange/services/2006/types}End').text
  response = element.find('{http://schemas.microsoft.com/exchange/services/2006/types}MyResponseType').text
  print_orgmode_entry(subject, start, end, location, response)
