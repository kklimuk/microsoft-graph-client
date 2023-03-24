# frozen_string_literal: true

require "httparty"
require "securerandom"
require_relative "graph/version"
require_relative "graph/batch"
require_relative "graph/body_formatter"

module Microsoft
  class Graph
    GRAPH_HOST = "https://graph.microsoft.com"
    BODY_METHODS = %w[POST PUT PATCH].freeze
    ALLOWED_METHODS = [*BODY_METHODS, "GET", "DELETE"].freeze

    def initialize(token: nil, error_handler: method(:error_handler), version: "1.0")
      @token = token
      @parser = URI::Parser.new
      @body_formatter = Microsoft::Graph::BodyFormatter.new
      @error_handler = error_handler
      @version = version
    end

    # stubs
    def get(*); end

    def post(*); end

    def put(*); end

    def patch(*); end

    def delete(*); end

    ALLOWED_METHODS.each do |method|
      define_method(method.downcase) { |*args, **kwargs| call(*args, method: method, **kwargs) }
    end

    def call(endpoint, token: @token, method: "GET", headers: {}, params: nil, body: nil, debug_output: nil)
      method = method.upcase
      raise ArgumentError, "`#{method}` is not a valid HTTP method." unless ALLOWED_METHODS.include?(method)

      url = URI.join(GRAPH_HOST, @parser.escape("v#{@version}/#{endpoint.gsub(%r{^/}, "")}"))
      headers = headers.merge(
        Authorization: "Bearer #{token}",
        Accept: "application/json"
      )
      headers[:"Content-Type"] = "application/json" if BODY_METHODS.include?(method)

      response = HTTParty.send(
        method.downcase,
        url,
        headers: headers,
        query: params,
        body: @body_formatter.call(body, method: method).to_json,
        parser: InstanceParser,
        debug_output: debug_output
      )

      case response.code
      when 200...400
        response.parsed_response
      when 400...600
        error = Error.new(
          "Received status code: #{response.code}. Check the `response` attribute for more details.",
          response.parsed_response,
          response.code
        )
        @error_handler.call(error)
      else
        raise "Unknown status code: #{response.code}"
      end
    end

    def batch(token: @token)
      batch = Batch.new(self, token: token)
      yield batch
      batch.call
    end

    def error_handler(error)
      raise error
    end

    class Error < StandardError
      attr_reader :response

      def initialize(message, response = nil, _code = nil)
        @response = response
        super message
      end
    end

    class InstanceParser < HTTParty::Parser
      def parse
        utf8_bom_in_body_encoding = UTF8_BOM.dup.force_encoding(@body.encoding)
        @body.gsub!(/\A#{utf8_bom_in_body_encoding}/, "").force_encoding("UTF-8") if @body.start_with?(utf8_bom_in_body_encoding)
        super
      end

      protected

      def json
        JSON.parse(body, object_class: JSONStruct)
      end
    end

    class JSONStruct < OpenStruct
      def self.format(key)
        HTTParty::Response.underscore(key.to_s).to_sym
      end

      def initialize(hash = {})
        super nil
        hash.each do |key, value|
          self[key] = value
        end
      end

      def []=(key, value)
        formatted_key = self.class.format(key)
        super formatted_key, value
      end
    end
  end
end
