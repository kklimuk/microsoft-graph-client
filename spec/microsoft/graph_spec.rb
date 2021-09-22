# frozen_string_literal: true

require "spec_helper"

RSpec.describe Microsoft::Graph do
  let(:graph) { described_class.new(token: token) }
  let(:token) { "token" }

  it "has a version number" do
    expect(Microsoft::Graph::VERSION).not_to be nil
  end

  describe "#call" do
    context "with a simple get request" do
      subject { graph.get("/me") }

      use_cassette "graph#call.get"

      it do
        is_expected.to have_attributes(
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
          display_name: "Kirill Klimuk",
          surname: "Klimuk",
          given_name: "Kirill",
          id: "89d5fafe0adc70ee",
          user_principal_name: "kklimuk@gmail.com"
        )
      end

      context "and the token is passed to the call directly" do
        subject { described_class.new.get("/me", token: token) }

        it do
          is_expected.to have_attributes(
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
            display_name: "Kirill Klimuk",
            surname: "Klimuk",
            given_name: "Kirill",
            id: "89d5fafe0adc70ee",
            user_principal_name: "kklimuk@gmail.com"
          )
        end
      end
    end

    context "with a request with a body" do
      subject do
        graph.patch("/me/drive/items/89D5FAFE0ADC70EE!106/workbook/worksheets/Sheet1/range(address='A56:B57')", body: {
          values: [%w[Hello 100], ["1/1/2016", nil]],
          formulas: [[nil, nil], [nil, "=B56*2"]],
          number_format: [[nil, nil], ["m-ddd", nil]]
        })
      end

      use_cassette "graph#call.body"

      it do
        is_expected.to have_attributes(
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#workbookRange",
          "@odata.type": "#microsoft.graph.workbookRange",
          "@odata.id": "/users('kklimuk%40gmail.com')/drive/items('89D5FAFE0ADC70EE%21106')/workbook/worksheets(%27%7B84FABE00-2D27-A843-B953-03E854DFA415%7D%27)/range(address=%27A56:B57%27)", # rubocop:disable Layout/LineLength
          address: "Sheet1!A56:B57",
          address_local: "Sheet1!A56:B57",
          column_count: 2,
          cell_count: 4,
          column_hidden: false,
          row_hidden: false,
          number_format: [
            %w[General General],
            %w[m-ddd General]
          ],
          column_index: 0,
          text: [
            %w[Hello 100],
            %w[1-Fri 200]
          ],
          formulas: [
            [
              "Hello",
              100
            ],
            [
              42_370,
              "=B56*2"
            ]
          ],
          formulas_local: [
            [
              "Hello",
              100
            ],
            [
              42_370,
              "=B56*2"
            ]
          ],
          formulas_r1c1: [
            [
              "Hello",
              100
            ],
            [
              42_370,
              "=R[-1]C*2"
            ]
          ],
          hidden: false,
          row_count: 2,
          row_index: 55,
          value_types: [
            %w[String Double],
            %w[Double Double]
          ],
          values: [
            [
              "Hello",
              100
            ],
            [
              42_370,
              200
            ]
          ]
        )
      end
    end
  end

  describe "#batch" do
    subject do
      graph.batch do |batch|
        batch.add("/me", id: "5e5108ca-b020-46e9-b557-9e13ec5b0781", method: "GET")
        batch.add(
          "/me/drive/items/89D5FAFE0ADC70EE!106/workbook/worksheets/Sheet1/range(address='A56:B57')",
          id: "75ac1a26-2210-4990-9ceb-fbc6f2b41e9e",
          method: "PATCH",
          body: {
            values: [%w[Hello 100], ["1/1/2016", nil]],
            formulas: [[nil, nil], [nil, "=B56*2"]],
            number_format: [[nil, nil], ["m-ddd", nil]]
          }
        )
      end
    end

    use_cassette "graph#batch"

    it do
      expect(subject.count).to eq(2)
    end
  end
end
