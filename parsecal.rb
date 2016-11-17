#!/usr/bin/ruby
require 'rubygems' # Unless you install from the tarball or zip.
require 'icalendar'
require 'date'
require 'open-uri'
include Icalendar # Probably do this in your class to limit namespace overlap


google_cal_url = 'TODO enter calender url here'
total = 0
min_by_month = Hash.new
min_by_tag = Hash.new
verbose = ARGV.index('-v')
open(google_cal_url) do |f|
	# Parse Calendar
	cals = Icalendar.parse(f)
	# Parser returns an array of calendars because a single file
	# can have multiple calendars.
	cal = cals.first

	cal.events.each do |event|
		#puts "#{event}"
		#puts "#{event.dtstart.to_time} -  #{event.dtend.to_time}"
		#puts "#{event.dtstart.to_s} -  #{event.dtend.to_s}"
		#puts "#{event.dtend.to_time - event.dtstart.to_time}"
	 #	puts "#{event.duration}"
		minutes = ((event.dtend.to_time - event.dtstart.to_time) / 60).to_i
		summary = event.summary.to_s
		month_key = "#{event.dtstart.year}-#{event.dtstart.month}"
		if min_by_month[month_key]
			min_by_month[month_key] = min_by_month[month_key] + minutes
		else
			min_by_month[month_key] = minutes
		end
		tag = summary[/(\s|\A)#(\w+)/]
		tag = :untagged if not tag
		if min_by_tag[tag]
			min_by_tag[tag] += minutes
		else
			min_by_tag[tag] = minutes
		end
		if verbose
			puts "#{event.dtstart.to_date.to_s}: #{minutes} min: " + summary
		end
		total += minutes
 	end
	puts "========================"
	puts "Total #{total/60}h #{total%60}min"
end

puts "========================"
min_by_month.each do |month,min|
	puts "#{month} #{min/60}h #{min%60}min"
end

puts "========================"
min_by_tag.each do |tag,min|
	puts "#{tag} #{min/60}h #{min%60}min"
end
