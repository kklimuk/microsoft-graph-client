# Microsoft::Graph

This is a [Microsoft Graph](https://docs.microsoft.com/en-us/graph/overview) client gem, since the official client library is [not supported](https://github.com/microsoftgraph/msgraph-sdk-ruby).

## Installation

Add this line to your application's Gemfile:

```ruby
gem "microsoft-graph-client"
```

And then execute:

    $ bundle install

Or install it yourself as:

    $ gem install microsoft-graph

## Usage

The gem supports both individual calls to the API. To issue either of these calls, it's important to initialize the Graph 
with the access token.

To get the authentication token, it's recommended to use the [Windows Azure Active Directory Authentication Library](https://github.com/AzureAD/azure-activedirectory-library-for-ruby) (ADAL) if you're building a console app.
You will have to [register the app](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-v2-netcore-daemon#register-and-download-the-app) and add permissions.

Here's some sample code to get you started:

```ruby
username      = 'admin@tenant.onmicrosoft.com'
password      = 'xxxxxxxxxxxx'
client_id     = 'xxxxx-xxxx-xxx-xxxxxx-xxxxxxx'
client_secret = 'xxxXXXxxXXXxxxXXXxxXXXXXXXXxxxxxx='
tenant        = 'tenant.onmicrosoft.com'
user_cred     = ADAL::UserCredential.new(username, password)
client_cred   = ADAL::ClientCredential.new(client_id, client_secret)
context       = ADAL::AuthenticationContext.new(ADAL::Authority::WORLD_WIDE_AUTHORITY, tenant)
resource      = "https://graph.microsoft.com"
tokens        = context.acquire_token_for_user(resource, client_cred, user_cred)

graph = Microsoft::Graph.new(token: tokens.access_token)
```

### Individual API Calls
To make individual calls, you can use the `#call` instance method on `Microsoft::Graph`:

`Microsoft::Graph#call(endpoint, method: "GET", headers: {}, params: nil, body: nil)`
- `endpoint` - the URL of the API without the version as a string. For example: `/me`.
- `method` - the preferred HTTP method as a string.
- `headers` - a hash of additional headers. `Microsoft::Graph` adds the appropriate JSON headers.
- `body` - the body, as a hash. The keys of the body will be converted from snake-case to camel-case when sent (i.e. `number_format` to `numberFormat`). 

As a syntactic sugar, `Microsoft::Graph` exposes the 5 HTTP methods as instance methods:
- `GET` - `#get`
- `POST`- `#post`
- `PUT`- `#put`
- `PATCH`- `#patch`
- `DELETE`- `#delete`

There are two additional quality of life features for Ruby
- You can pass the request body's keys in snake-case.
- The response will convert the API's camel-cased response to snake-case.

Below are a couple of examples:
```ruby
# Get a signed in user
# https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
graph.get("/me")
# alternatively
graph.call("/me", method: "GET")

# sample response - Microsoft::Graph converts the response keys to be snake-cased
# Microsoft::Graph::JSONStruct {
#   "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
#   display_name: "Kirill Klimuk",
#   surname: "Klimuk",
#   given_name: "Kirill",
#   id: "89d5fafe0adc70ee",
#   user_principal_name: "kklimuk@gmail.com" 
# }
```

```ruby
# update an excel worksheet
# https://docs.microsoft.com/en-us/graph/api/range-update?view=graph-rest-1.0&tabs=http
graph.patch("/me/drive/items/89D5FAFE0ADC70EA!106/workbook/worksheets/Sheet1/range(address='A56:B57')", body: {
  values: [%w[Hello 100], ["1/1/2016", nil]],
  formulas: [[nil, nil], [nil, "=B56*2"]],
  number_format: [[nil, nil], ["m-ddd", nil]] # in the API docs, this is described as numberFormat, but Microsoft::Graph allows you to use snake-cased keys
})
# alternatively
graph.call(
  "/me/drive/items/89D5FAFE0ADC70EA!106/workbook/worksheets/Sheet1/range(address='A56:B57')", 
  method: "PATCH",
  body: {
    values: [%w[Hello 100], ["1/1/2016", nil]],
    formulas: [[nil, nil], [nil, "=B56*2"]],
    number_format: [[nil, nil], ["m-ddd", nil]] # in the API docs, this is described as numberFormat, but Microsoft::Graph allows you to use snake-cased keys
  }
)

# sample response - Microsoft::Graph converts the response keys to be snake-cased
# Microsoft::Graph::JSONStruct {
#   "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#workbookRange",
#   "@odata.type": "#microsoft.graph.workbookRange",
#   "@odata.id": "/users('kklimuk%40gmail.com')/drive/items('89D5FAFE0ADC70EA%21106')/workbook/worksheets(%27%7B84FABE00-2D27-A843-B953-03E854DFA415%7D%27)/range(address=%27A56:B57%27)",
#   address: "Sheet1!A56:B57",
#   address_local: "Sheet1!A56:B57",
#   column_count: 2,
#   cell_count: 4,
#   column_hidden: false,
#     row_hidden: false,
#     number_format: [
#     %w[General General],
#     %w[m-ddd General]
#   ],
#   column_index: 0,
#   text: [
#     %w[Hello 100],
#     %w[1-Fri 200]
#   ],
#   formulas: [
#     [
#       "Hello",
#       100
#     ],
#     [
#       42_370,
#       "=B56*2"
#     ]
#   ],
#   formulas_local: [
#     [
#       "Hello",
#       100
#     ],
#     [
#       42_370,
#       "=B56*2"
#     ]
#   ],
#   formulas_r1c1: [
#     [
#       "Hello",
#       100
#     ],
#     [
#       42_370,
#       "=R[-1]C*2"
#     ]
#   ],
#   hidden: false,
#     row_count: 2,
#   row_index: 55,
#   value_types: [
#     %w[String Double],
#     %w[Double Double]
#   ],
#   values: [
#     [
#       "Hello",
#       100
#     ],
#     [
#       42_370,
#       200
#     ]
#   ]
# }
```

If an error occurs, an error will be thrown. You can inspect the error via its `response` method.

### Batched API Calls
Microsoft has rolled out a way to send the graph [multiple requests in one](https://docs.microsoft.com/en-us/graph/json-batching).
As a result, this client library supports it out of the box with the `#batch` instance method.

Here's an example that combines both of our individual calls into a single batched call.
```ruby
graph.batch do |batch|
    batch.add("/me", id: "abc", method: "GET")
    batch.add(
      "/me/drive/items/89D5FAFE0ADC70EA!106/workbook/worksheets/Sheet1/range(address='A56:B57')",
      id: "def",
      method: "PATCH",
      body: {
        values: [%w[Hello 100], ["1/1/2016", nil]],
        formulas: [[nil, nil], [nil, "=B56*2"]],
        number_format: [[nil, nil], ["m-ddd", nil]]
      }
    )
end
```

Please note: the current maximum number of calls in a single request is 20 according to Microsoft's docs.
The `#batch` method will group requests into batches of 20 to not violate this rule.

The output of the request will be an array of `Microsoft::Graph::Batch::Result`, which have both a `request` and `result` properties.
You can find the relevant result by searching for the request `id` you have passed in. 

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/kklimuk/microsoft-graph.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
