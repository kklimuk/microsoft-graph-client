# frozen_string_literal: true

module Microsoft
  class Graph
    class Batch
      attr_reader :requests

      def initialize(graph, token:, size: 20)
        @graph = graph
        @requests = []
        @token = token
        @size = size
      end

      def add(endpoint, id: SecureRandom.uuid, method: "GET", headers: {}, params: nil, body: nil, depends_on: nil)
        @requests << Request.new(
          endpoint,
          id: id,
          method: method,
          headers: headers,
          params: params,
          body: body,
          depends_on: depends_on
        )
      end

      def call
        @requests.each_slice(@size).flat_map do |group|
          requests_by_id = group.group_by(&:id).transform_values(&:first)
          group.first.depends_on = nil

          response = @graph.call("/$batch", method: "POST", token: @token, body: { requests: group.map(&:to_h) })
          response.responses.map do |current|
            Result.new(request: requests_by_id[current.id], response: current)
          end
        end
      end

      class Request
        def self.parser
          @parser ||= URI::Parser.new
        end

        def self.body_formatter
          @body_formatter ||= Microsoft::Graph::BodyFormatter.new
        end

        attr_reader :id, :endpoint, :method, :headers, :body
        attr_accessor :depends_on

        def initialize(endpoint, id:, method:, headers:, params:, body:, depends_on:)
          @id = id
          @endpoint = "#{self.class.parser.escape(endpoint)}#{params ? "?#{params.to_query}" : ""}"
          @method = method.upcase
          @headers = headers
          @headers[:Accept] = "application/json" if body
          @headers[:"Content-Type"] = "application/json" if body
          @body = body
          @depends_on = depends_on
        end

        def to_h
          hash = {
            id: @id,
            url: @endpoint,
            method: @method
          }
          hash[:headers] = @headers unless @headers.empty?
          hash[:body] = self.class.body_formatter.call(@body, method: @method) if @body
          hash[:dependsOn] = [@depends_on] if @depends_on

          hash
        end
      end

      class Result
        attr_reader :request, :response

        def initialize(request:, response:)
          @request = request
          @response = response
        end
      end
    end
  end
end
