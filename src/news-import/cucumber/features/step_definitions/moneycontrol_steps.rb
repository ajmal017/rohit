And(/^User search company (.*) and navigate to page$/) do |company|
  moneycontrol_page.search_company(company)
  moneycontrol_page.click_search_company
end

And(/^User click on news link$/) do
  moneycontrol_page.click_news_link
end

And(/^User get newsurl company (.*), symbol (.*)$/) do |company, symbol|
  moneycontrol_page.get_url_and_log(company, symbol)
end

And(/^User get news (.*)$/) do |symbol|
  moneycontrol_page.get_news_and_log(symbol)
end

And(/^Get news from link$/) do
  moneycontrol_page.get_news_and_log(String(ENV['link']).split('@')[0])
end

