require "google/apis/drive_v3"
require "google/apis/slides_v1"
require "csv"

class Slide

  def initialize(user, template_presentation_id)
    @user = user
    @template_presentation_id = template_presentation_id
  end

  def create_for_all_keywords
    today_string = Date.today.strftime("%Y%m%d")
    current_month = Time.now.month
    total_days = Time.days_in_month(current_month, Time.now.year)
    csv_collection = CSV.read("#{Rails.root}/storage/result/summary/result_after_screenshot_#{today_string}.csv", headers: true)
    google_authorization = GoogleAuthorization.new(@user).authorize
    drive_service = ::Google::Apis::DriveV3::DriveService.new
    drive_service.authorization = google_authorization
    slides_service = Google::Apis::SlidesV1::SlidesService.new
    slides_service.authorization = google_authorization
    file = Google::Apis::DriveV3::File.new
    file.name = "Brand Control #{today_string}"

    drive_response = drive_service.copy_file(@template_presentation_id, file)
    presentation_id = drive_response.id
    presentation = slides_service.get_presentation(presentation_id)
    summary_table_template_page_id = presentation.slides[0].object_id_prop
    summary_template_page_id = presentation.slides[1].object_id_prop
    not_found_template_page_id = presentation.slides[2].object_id_prop
    found_template_page_id = presentation.slides[3].object_id_prop

    # add summary page
    requests = []
    google_found_keywords_count = csv_collection.count { |row| row["google_total_found_in_current_month_count"].to_i > 0 }
    requests << {
      replace_all_text: {
        replace_text: "#{google_found_keywords_count}",
        page_object_ids: [summary_template_page_id],
        contains_text: {
          text: "{{google_found_keywords_count}}"
        }
      }
    }

    yahoo_found_keywords_count = csv_collection.count { |row| row["yahoo_total_found_in_current_month_count"].to_i > 0 }
    requests << {
      replace_all_text: {
        replace_text: "#{yahoo_found_keywords_count}",
        page_object_ids: [summary_template_page_id],
        contains_text: {
          text: "{{yahoo_found_keywords_count}}"
        }
      }
    }

    req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
    response = slides_service.batch_update_presentation(presentation_id, req)

    csv_collection.each_slice(20).reverse_each do |group|
      requests = []
      requests << {
        duplicate_object: {
          object_id_prop: summary_table_template_page_id
        }
      }

      req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
      response = slides_service.batch_update_presentation(presentation_id, req)
      summary_table_page_id = response.replies[0].duplicate_object.object_id_prop
      summary_table_page = slides_service.get_presentation_page(presentation_id, summary_table_page_id)
      summary_table_element = summary_table_page.page_elements.detect { |e| e.table }

      group.reverse_each.with_index do |row, index|
        keyword = row["name"]
        google_found_count = row["google_found_count"].to_i
        yahoo_found_count = row["yahoo_found_count"].to_i
        google_total_found_count = row["google_total_found_in_current_month_count"].to_i
        yahoo_total_found_count = row["yahoo_total_found_in_current_month_count"].to_i

        # add summary table page
        requests = []
        requests << {
          replace_all_text: {
            replace_text: "#{current_month}",
            page_object_ids: [summary_table_page.object_id_prop],
            contains_text: {
              text: "{{current_month}}"
            }
          }
        }

        requests << {
          insert_table_rows: {
            table_object_id: summary_table_element.object_id_prop,
            insert_below: true,
            number: 1,
          }
        }

        requests << {
          insert_text: {
            object_id_prop: summary_table_element.object_id_prop,
            cell_location: {
              row_index: 1,
              column_index: 0,
            },
            text: keyword
          }
        }

        google_found_text = "×"
        if google_total_found_count >= 1
          if google_found_count == 0
            google_found_text = "※(#{google_total_found_count}日)"
          else
            google_found_text = "○(#{google_total_found_count}日)"
          end
        end

        requests << {
          insert_text: {
            object_id_prop: summary_table_element.object_id_prop,
            cell_location: {
              row_index: 1,
              column_index: 1,
            },
            text: google_found_text
          }
        }

        yahoo_found_text = "×"
        if yahoo_total_found_count >= 1
          if yahoo_found_count == 0
            yahoo_found_text = "※(#{yahoo_total_found_count}日)"
          else
            yahoo_found_text = "○(#{yahoo_total_found_count}日)"
          end
        end

        requests << {
          insert_text: {
            object_id_prop: summary_table_element.object_id_prop,
            cell_location: {
              row_index: 1,
              column_index: 2,
            },
            text: yahoo_found_text
          }
        }

        requests << {
          insert_text: {
            object_id_prop: summary_table_element.object_id_prop,
            cell_location: {
              row_index: 1,
              column_index: 3,
            },
            text: yahoo_found_text
          }
        }

        req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
        response = slides_service.batch_update_presentation(presentation_id, req)
      end
    end

    csv_collection.reverse_each do |row|
      keyword = row["name"]
      google_found_count = row["google_found_count"].to_i
      yahoo_found_count = row["yahoo_found_count"].to_i
      google_total_found_count = row["google_total_found_in_current_month_count"].to_i
      yahoo_total_found_count = row["yahoo_total_found_in_current_month_count"].to_i
      google_blob = ActiveStorage::Blob.where(filename: "#{keyword}_google_#{today_string}.png").order(id: :desc).first
      yahoo_blob = ActiveStorage::Blob.where(filename: "#{keyword}_yahoo_#{today_string}.png").order(id: :desc).first
      google_image_url =  google_blob.service_url
      yahoo_image_url = yahoo_blob.service_url

      if google_found_count == 0 && yahoo_found_count == 0
        template_page_id = not_found_template_page_id
      else
        template_page_id = found_template_page_id
      end

      requests = []
      requests << {
        duplicate_object: {
          object_id_prop: template_page_id
        }
      }

      req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
      response = slides_service.batch_update_presentation(presentation_id, req)
      page_id = response.replies[0].duplicate_object.object_id_prop

      requests = []
      requests << {
        replace_all_text: {
          replace_text: keyword,
          page_object_ids: [page_id],
          contains_text: {
            text: "{{keyword}}"
          }
        }
      }

      if google_found_count >= 1 || yahoo_found_count >= 1
        if google_found_count >= 1 && yahoo_found_count >= 1
          google_and_yahoo_text = "GoogleサジェストとYahoo!"
        elsif google_found_count >= 1
          google_and_yahoo_text = "Googleサジェスト"
        elsif yahoo_found_count >= 1
          google_and_yahoo_text = "Yahoo!サジェスト"
        end

        requests << {
          replace_all_text: {
            replace_text: google_and_yahoo_text,
            page_object_ids: [page_id],
            contains_text: {
              text: "{{google_and_yahoo_text}}"
            }
          }
        }

        if google_found_count >= 1
          google_summary_text = "Googleサジェスト　#{current_month}月表示日数：#{google_total_found_count}日/#{total_days}日"
        else
          google_summary_text = ""
        end

        requests << {
          replace_all_text: {
            replace_text: google_summary_text,
            page_object_ids: [page_id],
            contains_text: {
              text: "{{google_summary_text}}"
            }
          }
        }

        if yahoo_found_count >= 1
          yahoo_summary_text = "Yahoo!サジェスト　#{current_month}月表示日数：#{yahoo_total_found_count}日/#{total_days}日"
        else
          yahoo_summary_text = ""
        end

        requests << {
          replace_all_text: {
            replace_text: yahoo_summary_text,
            page_object_ids: [page_id],
            contains_text: {
              text: "{{yahoo_summary_text}}"
            }
          }
        }
      end

      requests << {
        replace_all_shapes_with_image: {
          image_replace_method: "CENTER_INSIDE",
          page_object_ids: [page_id],
          contains_text: {
            text: "{{google_image}}"
          },
          image_url: google_image_url
        }
      }

      requests << {
        replace_all_shapes_with_image: {
          image_replace_method: "CENTER_INSIDE",
          page_object_ids: [page_id],
          contains_text: {
            text: "{{yahoo_image}}"
          },
          image_url: yahoo_image_url
        }
      }

      req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
      response = slides_service.batch_update_presentation(
        presentation_id,
        req
      )
      sleep 1
    end

    requests = []
    requests << {
      delete_object: {
        object_id_prop: found_template_page_id,
      }
    }

    requests << {
      delete_object: {
        object_id_prop: not_found_template_page_id,
      }
    }

    requests << {
      delete_object: {
        object_id_prop: summary_table_template_page_id,
      }
    }

    req = Google::Apis::SlidesV1::BatchUpdatePresentationRequest.new(requests: requests)
    response = slides_service.batch_update_presentation(presentation_id, req)
  end
end
