import argparse
import datetime
import re

# initiate argument parser
import uuid

parser = argparse.ArgumentParser(
    description='convert an outlook calendar file (.hol) to an icalendar file (.ics)')

parser.add_argument('source_filename',
                    help='outlook calendar source file (.hol)')

parser.add_argument('--dest', dest='destination_filename',
                    default=None,
                    help='optional destination filename (.ics)')

args = parser.parse_args()

if not args.source_filename.lower().endswith("hol"):
    raise ValueError("Source file has to end with the hol extension")

source_filename = args.source_filename

if args.destination_filename is not None:
    if not args.destination_filename.lower().endswith("ics"):
        raise ValueError("Destination file has to end with the ics extension")
    destination_filename = args.destination_filename
else:
    destination_filename = source_filename.split(".")[0] + ".ics"

print(f"Attempting to convert {source_filename} into {destination_filename}")

def line_to_event_tuple(line):
    day_title, date_str = line.split(",")
    # hol does not specify timezone, I think, so we default to the local timezone of the user
    begin_datetime = datetime.datetime.strptime(date_str.strip(), "%Y/%m/%d")#.astimezone()
    # print(day_title, date_str, begin_datetime, begin_datetime.tzinfo)
    return day_title, begin_datetime


ics_datetime_format = "%Y%m%dT%H%M%S"


def write_ics_file(events, destination_filename, title):
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//{title}//EN",
    ]

    for name, start_date in events:
        creation_time = datetime.datetime.now().astimezone().strftime(ics_datetime_format)
        start_date = start_date.strftime("%Y%m%d")
        event_lines = [
            "BEGIN:VEVENT",
            f"UID:{uuid.uuid4()}",
            f"DTSTAMP:{creation_time}",
            f"DTSTART;VALUE=DATE:{start_date}",
            f"SUMMARY:{name}",
            "END:VEVENT",
            ]
        lines += event_lines

    lines.append("END:VCALENDAR")
    # the ics format must have a CRLF at the end of each line
    # see https://icalendar.org/iCalendar-RFC-5545/3-1-content-lines.html
    ics_str = "\r\n".join(lines)

    # finally, write to the new file
    with open(destination_filename, 'w') as out_file:
        out_file.writelines(ics_str)


# TODO: there can be several section in a HOL file - the header and the count
#  will tell us what to do - but do this later!
header_pattern = r"\[(?P<title>.+)\]\s*(?P<count>[0-9]+)"

with open(source_filename, encoding='utf-16') as handler:
    header = handler.readline()
    # print(header)
    matches = re.match(header_pattern, header)

    if matches:
        title = matches.group('title')
        count = matches.group('count')
        # print(f"title={title}, count={count}")

    # how do add the title to the new calendar?

    # read the rest of the rows
    events = map(line_to_event_tuple, handler.readlines())
    write_ics_file(events, destination_filename, title)

    print("You can try to validate the resultant file using this webform: https://icalendar.org/validator.html")



