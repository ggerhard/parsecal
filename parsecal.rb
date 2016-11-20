#!/usr/bin/ruby
require 'rubygems' # Unless you install from the tarball or zip.
require 'icalendar'
require 'date'
require 'open-uri'
require 'axlsx'
include Icalendar # Probably do this in your class to limit namespace overlap

# set calendar URL here
google_cal_url = 'TODO Add your calendar URL here'

# class for excel sheets
class MonthSheet
	attr_accessor :sheet, :title , :headline_style, :time_style, :date_style, :cell_style, :sum_style

	def initialize(title)
		@title = title
		@sheet = $wb.add_worksheet(:name => title)
		# initialize styles
		@headline_style = sheet.styles.add_style(:format_code => "HH:MM", :border => Axlsx::STYLE_THIN_BORDER, :b => true)
		@time_style = sheet.styles.add_style(:format_code => "HH:MM", :border => Axlsx::STYLE_THIN_BORDER)
		@date_style = sheet.styles.add_style(:format_code => "dd.mm.yyyy", :border => Axlsx::STYLE_THIN_BORDER)
		@cell_style = sheet.styles.add_style(:border => Axlsx::STYLE_THIN_BORDER)
		@sum_style = sheet.styles.add_style(:format_code => "[HH]:MM",:b => true ,:border => Axlsx::STYLE_THIN_BORDER)
		# add title row
		@sheet.add_row ["Date", "From", "To", "Duration", "Task"], :style => headline_style
	end

	def add_row
		return @sheet.add_row
	end

	def add_sum(formula)
		return @sheet.add_row ["", "", "", formula, ""], :style => @sum_style
	end

end

# initialize variables
total = 0
min_by_month = Hash.new
min_by_tag = Hash.new
verbose = ARGV.index('-v')
first_duration = nil
last_duration = nil
sheet = nil

# initialize excel worksheet
xls_package = Axlsx::Package.new
$wb = xls_package.workbook

# parse Calendar and create excel sheets
open(google_cal_url) do |f|
	# Parse Calendar
	cals = Icalendar::Calendar.parse(f)
	# Parser returns an array of calendars because a single file
	# can have multiple calendars.
	cal = cals.first

	cal.events.each do |event|
		#puts "#{event}"
		#puts "#{event.dtstart.to_time} -  #{event.dtend.to_time}"
		#puts "#{event.dtstart.to_s} -  #{event.dtend.to_s}"
		#puts "#{event.dtend.to_time - event.dtstart.to_time}"
	 	#puts "#{event.duration}"

		minutes = ((event.dtend.to_time - event.dtstart.to_time) / 60).to_i
		summary = event.summary.to_s

		month_key = "#{event.dtstart.year}-#{event.dtstart.month}"
		if min_by_month[month_key]
			min_by_month[month_key] = min_by_month[month_key] + minutes
		else
			min_by_month[month_key] = minutes
			sheet.add_sum "=SUM(#{first_duration.r}:#{last_duration.r})" if sheet != nil
			sheet = MonthSheet.new(month_key)
		end

		# output to spreadsheet
		rw = sheet.add_row
		c_date = rw.add_cell(event.dtstart.to_time)
		c_date.style = sheet.date_style
		c_start = rw.add_cell(event.dtstart.to_time)
		c_start.style = sheet.time_style
		c_end = rw.add_cell(event.dtend.to_time)
		c_end.style = sheet.time_style
		c_diff = rw.add_cell("= #{c_end.r}-#{c_start.r}")
		c_diff.style = sheet.time_style
		c_sum = rw.add_cell(summary)
		c_sum.style = sheet.cell_style


		first_duration = c_diff if first_duration.nil?
		last_duration = c_diff

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
	sheet.add_sum "=SUM(#{first_duration.r}:#{last_duration.r})"
	puts "Total #{total/60}h #{total%60}min"
end

sheet2 = $wb.add_worksheet(:name => "Per Tag")


min_by_tag.each do |tag,min|
	sheet2.add_row ["#{tag}", "#{min/60}h #{min%60}min"]
	# puts "#{tag} #{min/60}h #{min%60}min"
end

xls_package.serialize('timesheet.xlsx')
