# frozen_string_literal: true

module Microsoft
  class Graph
    class BodyFormatter
      def call(body, method:)
        return nil unless Microsoft::Graph::BODY_METHODS.include?(method)
        return nil unless body

        body.transform_keys(&method(:camelize))
      end

      private

      def camelize(key)
        string = key.to_s
        return string unless string.include?("_")

        string = string.sub(/^(?:(?=\b|[A-Z_])|\w)/, &:downcase)
        string.gsub(%r{(?:_|(/))([a-z\d]*)}) do
          "#{Regexp.last_match(1)}#{Regexp.last_match(2).capitalize}"
        end.gsub("/", "::")
      end
    end
  end
end
