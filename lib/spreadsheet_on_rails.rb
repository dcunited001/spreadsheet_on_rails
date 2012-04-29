module SpreadsheetOnRails

  def self.render_xls_string(spreadsheet)
<<RENDER
    workbook = Spreadsheet::Workbook.new
    #{spreadsheet}
    blob = StringIO.new("")
    workbook.write(blob)
    blob.string
RENDER
  end

  # render csv: xls_filename, locals: {:csv => true}
  def self.render_csv_string(spreadsheet)
<<RENDER
    workbook = Spreadsheet::Workbook.new
    #{spreadsheet}
    default_sheet = workbook.worksheets.first
    CSV.generate do |c|
      default_sheet.each do |r|
        c << r.to_a
      end
    end
RENDER
  end

end

# Setups the template handling
require "action_view/template"
require 'spreadsheet'
ActionView::Template.register_template_handler :rxls, lambda { |template|
  unless template.locals.include?(:csv)
    SpreadsheetOnRails.render_xls_string(template.source)
  else
    SpreadsheetOnRails.render_csv_string(template.source)
  end
}

# Why doesn't the aboce template handler catch this one as well?
# Added for backwards compatibility.
ActionView::Template.register_template_handler :"xls.rxls", lambda { |template|
  unless template.locals.include?(:csv)
    SpreadsheetOnRails.render_xls_string(template.source)
  else
    SpreadsheetOnRails.render_csv_string(template.source)
  end
}

# Adds support for `format.xls`
require "action_controller"
Mime::Type.register "application/xls", :xls
Mime::Type.register "text/csv", :csv

ActionController::Renderers.add :xls do |filename, options|
  send_data(render_to_string(options), :filename => "#{filename}.xls", :type => "application/xls", :disposition => "attachment")
end

ActionController::Renderers.add :csv do |filename, options|
  send_data(render_to_string(options), :filename => "#{filename}.csv", :type => "text/csv", :disposition => "attachment")
end
