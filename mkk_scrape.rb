require "mechanize"
require "active_support"
require "active_support/core_ext"
require "parallel"
require "pry"

def modify_link(link)
  link = link[0] == "/" ? "http://www.mkk.gov.kg#{link}" : link
  return link
end

a = Mechanize.new do |agent|
  agent.user_agent_alias = "Mac Safari"
  agent.open_timeout = 50
  agent.read_timeout = 50
  agent.request_headers = {
    "Host" => "mkk.gov.kg",
    "User-Agent" => "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:57.0) Gecko/20100101 Firefox/57.0",
    "Accept" => "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language" => "en-US,en;q=0.5",
    "Referer" => "http://www.mkk.gov.kg/contents/view/id/664/pid/157",
    "Connection" => "keep-alive",
    "Upgrade-Insecure-Requests" => "1",
    "Cookie" => "PHPSESSID=qvcaaa9a0b7sogs19v8u9b85n4; _ga=GA1.3.1094423778.1516545618; _gid=GA1.3.2003521789.1516545618",
    "Cache-Control" => "max-age=0",
  }
end

page = a.get "http://www.mkk.gov.kg/contents/view/id/615/pid/95"
pages = page.links.select { |l| l.href.include?("157") }
Parallel.each(pages, in_threads: 10, :progress => "doing") do |p|
  # puts "processed #{index} out of #{pages.count}"
  link = modify_link(p.href)
  page = a.get link
  page.links.select { |l| l.href.include?("xls") }.each do |file|
    file_name = File.basename(file.href)
    next if File.exist?(file_name)
    begin
      a.get(modify_link(file.href)).save(file_name)
    rescue => ex
      puts "ERROR:#{ex.inspect} link:#{file.href}"
    end
  end
end
