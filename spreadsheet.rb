#
#  Google Spreadsheet API module
#    http://www.binzume.net/

require "rexml/document"
require_relative 'httpclient'

class GSession
  def initialize
    @client = HTTPClient.new
    @auth = nil
  end

  def login account
    r = @client.post('https://www.google.com/accounts/ClientLogin',account)
    if r.body =~ /Auth=([^\s]+)/
      @auth = $1
      return true
    end
    return false
  end

  def get url
    return @client.get(url, {"Authorization"=>"GoogleLogin auth=#{@auth}"})
  end
  def post url,data,headers=nil
    return @client.post(url, data, {"Authorization"=>"GoogleLogin auth=#{@auth}", "Content-Type" => "application/atom+xml"})
  end
  def put url,data,headers=nil
    return @client.put(url, data, {"Authorization"=>"GoogleLogin auth=#{@auth}", "Content-Type" => "application/atom+xml"})
  end
end

class GWorksheet
  attr_accessor :url
  def initialize session, url,cells_url, title
    @session = session
    @url = url
    @cells_url = cells_url
    @title = title
    @row_count = 0
    @col_count = 0
    @edit_url = nil
    @loaded = false
  end

  def load_meta
    r = @session.get(@url)
    # p r.body
    doc = REXML::Document.new(r.body)
    @title = doc.elements['//entry/title'].text
    @row_count = doc.elements['//entry/gs:rowCount'].text.to_i
    @col_count = doc.elements['//entry/gs:colCount'].text.to_i
    doc.elements.each('//entry/link'){|e|
      if e.attributes['rel'] == "edit"
        @edit_url = e.attributes['href']
      end
    }
    @loaded = true
  end

  def update_meta
    data = <<-EOD
      <entry xmlns='http://www.w3.org/2005/Atom'
             xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
        <title>#{@title}</title>
        <gs:rowCount>#{@row_count}</gs:rowCount>
        <gs:colCount>#{@col_count}</gs:colCount>
      </entry>
    EOD
    r = @session.put(@edit_url, data)
  end

  def set_cell row,col,value
    cells = load(row,col,row,col)
    id = cells[[row,col]].elements['id'].text

    data = <<-EOD
      <entry xmlns="http://www.w3.org/2005/Atom"
          xmlns:gs="http://schemas.google.com/spreadsheets/2006">
        <id>#{id}</id>
        <link rel="edit" type="application/atom+xml"
          href="#{id}"/>
        <gs:cell row="#{row}" col="#{col}" inputValue="#{value}"/>
      </entry>
    EOD
    r = @session.put(id, data)
  end

  def row_count
    unless @loaded
      load_meta
    end
    @row_count
  end

  def col_count
    unless @loaded
      load_meta
    end
    @col_count
  end

  def title
    unless @loaded
      load_meta
    end
    @title
  end

  def load min_row,min_col, max_row,max_col
    url = @cells_url + "?return-empty=true&min-row=#{min_row}&max-row=#{max_row}&min-col=#{min_col}&max-col=#{max_col}"
    r = @session.get(url)
    #p r.body
    doc = REXML::Document.new(r.body)

    cells = {}
    doc.elements.each('//feed/entry'){|entry|
      cell = entry.elements['gs:cell']
      col = cell.attributes['col'].to_i
      row = cell.attributes['row'].to_i
      cells[[row,col]] = entry
      #p [row,col]
    }
    #p [min_row,min_col, max_row,max_col]

    return cells
  end

  def load_all
    unless @loaded
      load_meta
    end
    @cells = load(1, 1, @row_count, @col_count)
  end

  def [] row, col = nil
    if col == nil
      cells = load(row,1,row,col_count)
      return cells.map{|k,v| v.elements['gs:cell'].text}
    end

    if @cells && @cells[[row,col]]
      cells = @cells
    else
      cells = load(row,col,row,col)
    end
    return cells[[row,col]].elements['gs:cell'].text
  end

  def []= row,col,value
    values = {[row,col]=>value}
    batch values
  end

  def << rowdata
    values = {}
    row = row_count + 1
    col = 1
    rowdata.each{|value|
      values[[row,col]] = value
      col += 1
    }
    batch values
  end

  def batch values
    unless @loaded
      load_meta
    end
    max_row = 0
    max_col = 0
    min_row = 1000000000
    min_col = 1000000000

    values.each{|rowcol,value|
      row = rowcol[0]
      col = rowcol[1]
      max_row = [max_row,row].max
      max_col = [max_col,col].max
      min_row = [min_row,row].min
      min_col = [min_col,col].min
    }

    if max_row > @row_count || max_col > @col_count
      @row_count = [@row_count, max_row].max
      @col_count = [@col_count, max_col].max
      update_meta
    end

    cells = load(min_row, min_col, max_row, max_col)

    xml = <<-EOD
      <feed xmlns="http://www.w3.org/2005/Atom"
            xmlns:batch="http://schemas.google.com/gdata/batch"
            xmlns:gs="http://schemas.google.com/spreadsheets/2006">
        <id>#{@cells_url}</id>
    EOD

    values.each{|rowcol,value|
      row = rowcol[0]
      col = rowcol[1]
      edit_url= nil
      id = cells[[row,col]].elements['id'].text
      cells[[row,col]].each_element('link') {|e|
        if e.attributes['rel'] == "edit"
          edit_url = e.attributes['href']
        end
      }

      cellxml = REXML::Element.new("gs:cell")
      cellxml.attributes["row"] = row;
      cellxml.attributes["col"] = col;
      cellxml.attributes["inputValue"] = value;

      xml << <<-EOD
        <entry>
          <batch:id>#{row}_#{col}</batch:id>
          <batch:operation type="update"/>
          <id>#{id}</id>
          <link rel="edit" type="application/atom+xml"
            href="#{edit_url}"/>
          #{cellxml.to_s}
        </entry>
      EOD
    }

    xml << <<-EOD
      </feed>
    EOD
    r = @session.post(@cells_url+"/batch",xml)

    load_meta
  end


  def set_cells row,col,values
    cells = {}
    scol = col
    values.each{|r|
      r.each{|value|
        cells[[row,col]]=value
        col+=1
      }
      row+=1
      col = scol
    }
    batch cells
  end

  def save
  end

end

class GSpreadsheet

  def initialize session, url
    @session = session
    if url !~ /https?:/
      url = "https://spreadsheets.google.com/feeds/worksheets/#{url}/private/full"
    end
    @url = url
    @worksheets = []
    @loaded = false
  end
  def load
    r = @session.get(@url)
    # puts r.body
    doc = REXML::Document.new(r.body)
    @id = doc.elements['//feed/id'].text
    @title = doc.elements['//feed/title'].text
    doc.elements.each('//feed/entry'){|entry|
      cells_url = entry.elements['id'].text
      entry.elements.each('link'){|e|
        if e.attributes['rel'] == "http://schemas.google.com/spreadsheets/2006#cellsfeed"
          cells_url = e.attributes['href']
        end
      }

      ws = GWorksheet.new(@session, entry.elements['id'].text ,cells_url, entry.elements['title'].text)
      @worksheets.push << ws
    }
    @loaded = true
  end
  def worksheets
    unless @loaded
      load
    end
    return @worksheets;
  end

  def title
    unless @loaded
      load
    end
    @title
  end
end




# google-spreadsheet-ruby compatible interface (see https://github.com/gimite/google-spreadsheet-ruby )
module GoogleSpreadsheet
  def self.login email,passwd
    session = GSession.new
    def session.spreadsheet_by_key spreadsheet_key
      GSpreadsheet.new(self, spreadsheet_key)
    end
    account = {
      'accountType' => 'HOSTED_OR_GOOGLE',
      'Email' =>  email,
      'Passwd' => passwd,
      'service' => 'wise', # GAE:ah CAL:cl PLUS:oz SPREADSHEET:wise
      'source' => 'simple-spreadsheet.binzume.net',
    }

    unless session.login(account)
      return nil
    end
    return session
  end
end
