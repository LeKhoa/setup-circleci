require "ferrum"
require "googleauth"
require "google/apis/sheets_v4"
require "csv"
require "fileutils"

class Screenshot

  SEARCH_SERVICES = {
    google: {
      name: "Google",
      url: "https://google.com",
      search_input_name: "q",
      selector: ".tsf",
      item_node: ".tsf ul li",
    },
    yahoo: {
      name: "Yahoo!",
      url: "https://search.yahoo.co.jp/",
      search_input_name: "p",
      selector: "#SaC",
      item_node: "ul#sugres li.suggest",
    },
  }

  def initialize(keywords, included_keywords)
    @keywords = keywords
    @included_keywords = included_keywords
  end

  def call
    positive_highlight_css = "{ background-color: #CCF5CB; border-style: dashed; border-color: #4DCA7F; border-width: 1px }"
    negative_highlight_css = "{ background-color: #FEF4F4; border-style: dashed; border-color: #FAA051; border-width: 1px }"
    current_highlight_css = negative_highlight_css
    browser = Ferrum::Browser.new
    user_agent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.120 Safari/537.36"
    browser.headers.set({ "User-Agent" => user_agent })

    today_string = Date.today.strftime("%Y%m%d")
    yesterday_string = Date.today.prev_day.strftime("%Y%m%d")
    yesterday_rows = CSV.read("result/summary/result_after_screenshot_#{yesterday_string}.csv", headers: true) rescue []

    dirname = "#{Rails.root}/storage/result/summary"
    FileUtils.mkdir_p(dirname) unless File.directory?(dirname)

    CSV.open("#{Rails.root}/storage/result/summary/result_after_screenshot_#{today_string}.csv", "w", { force_quotes: true }) do |csv|
      headers = ["name", "google_found_count", "google_total_found_in_current_month_count", "yahoo_found_count", "yahoo_total_found_in_current_month_count"]
      csv << headers

      @keywords.each do |keyword|
        row = [keyword]

        SEARCH_SERVICES.each do |search_service, options|
          found_count = 0
          search_input_name = options[:search_input_name]
          item_node = options[:item_node]
          selector = options[:selector]

          begin
            url = "#{options[:url]}/search?#{search_input_name}=#{keyword}&hl=ja"
            browser.goto(url)
            search_input = browser.at_css("input[name='#{search_input_name}']")
            search_input.focus.type(" ")
            sleep 1

            nodes = browser.css(item_node)
            nodes.each_with_index do |node, index|
              if @included_keywords.any? { |kw| node.text.include?(kw) }
                highlight_style = "#{item_node}:nth-child(#{index + 1}) #{current_highlight_css}"
                browser.add_style_tag(content: highlight_style)
                found_count = 1
              end
            end

            if search_service == :yahoo
              browser.execute(%(s2=document.getElementById("Sb_2");s1=document.getElementById("Si1");s2.append(s1)));
              browser.execute(%(e=document.getElementById("SaA");e.parentNode.removeChild(e)));
              browser.execute(%(e=document.getElementsByClassName("_pdd")[0];e.parentNode.removeChild(e)));
              browser.execute(%(e=document.getElementsByClassName("searchForm-opt")[0];e.parentNode.removeChild(e)));
              browser.add_style_tag(content: ".searchForm .sbox_1 { min-width: fit-content }")
              height = browser.evaluate(%(document.getElementById('Sb_2').offsetHeight)) + 20
              browser.add_style_tag(content: "#SaC { min-height: #{height}px }")

              item_node = "#Si1 a"
              nodes = browser.css(item_node)
              nodes.each_with_index do |node, index|
                if @included_keywords.any? { |kw| node.text.include?(kw) }
                  highlight_style = "#{item_node}:nth-child(#{index + 1}) #{current_highlight_css}"
                  browser.add_style_tag(content: highlight_style)
                end
              end
            end

            if search_service == :google
              browser.add_style_tag(content: "#tsf { width: 786px }")
            end

            sleep 1
            file_name = "#{keyword}_#{search_service}_#{today_string}.png"
            path = "result/#{keyword}/#{search_service}/#{keyword}_#{search_service}_#{today_string}.png"
            dirname = File.dirname(path)
            FileUtils.mkdir_p(dirname) unless File.directory?(dirname)
            browser.screenshot(path: path, selector: selector)
            # base64_image = browser.screenshot(selector: selector, encoding: :base64)
            # io_image = StringIO.new(base64_image)
            ActiveStorage::Blob.create_after_upload!(
              io: File.open(path),
              filename: file_name,
              content_type: "image/png",
            )
            puts "#{search_service} captured for #{keyword}"
          end

          row << found_count
          keyword_row = yesterday_rows.detect { |r| r["name"] == keyword }

          total_count = 0

          if keyword_row
            total_count = keyword_row["#{search_service}_total_found_in_current_month_count"].to_i
          end

          total_count += 1 if found_count >= 1

          row << total_count
        end

        csv << row
      end
    end

    browser.quit
  end
end
