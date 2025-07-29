# frozen_string_literal: true

module Axlsx
  # The Cell Serializer class contains the logic for serializing cells based on their type.
  class CellSerializer
    class << self
      # Calls the proper serialization method based on type.
      # @param [Integer] row_index The index of the cell's row
      # @param [Integer] column_index The index of the cell's column
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def to_xml_string(row_index, column_index, cell, str = +'')
        str << '<c r="'
        str << Axlsx.col_ref(column_index) << Axlsx.row_ref(row_index)
        str << '" s="' << cell.style_str << '" '
        return str << '/>' if cell.value.nil?

        method = cell.type
        send(method, cell, str)
        str << '</c>'
      end

      # builds an xml text run based on this cells attributes.
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def run_xml_string(cell, str = +'')
        if cell.is_text_run?
          valid = RichTextRun::INLINE_STYLES - [:value, :type]
          data = Axlsx.instance_values_for(cell).transform_keys(&:to_sym)
          data = data.select { |key, value| !value.nil? && valid.include?(key) }
          RichText.new(cell.value.to_s, data).to_xml_string(str)
        elsif cell.contains_rich_text?
          cell.value.to_xml_string(str)
        else
          str << '<t>' << cell.clean_value << '</t>'
        end
        nil
      end

      # serializes cells that are type iso_8601
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def iso_8601(cell, str = +'')
        value_serialization 'd', cell.value, str
      end

      # serializes cells that are type date
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def date(cell, str = +'')
        value_serialization false, DateTimeConverter.date_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type time
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def time(cell, str = +'')
        value_serialization false, DateTimeConverter.time_to_serial(cell.value).to_s, str
      end

      # Serializes cells that are type boolean
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def boolean(cell, str = +'')
        value_serialization 'b', cell.value.to_s, str
      end

      # Serializes cells that are type float
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def float(cell, str = +'')
        numeric cell, str
      end

      # Serializes cells that are type integer
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def integer(cell, str = +'')
        numeric cell, str
      end

      # Serializes cells that are type formula
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def formula_serialization(cell, str = +'')
        str << 't="str"><f>' << cell.clean_value.delete_prefix(FORMULA_PREFIX) << '</f>'
        str << '<v>' << cell.formula_value.to_s << '</v>' unless cell.formula_value.nil?
      end

      # Serializes cells that are type array formula
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def array_formula_serialization(cell, str = +'')
        str << 't="str">' << '<f t="array" ref="' << cell.r << '">' << cell.clean_value.delete_prefix(ARRAY_FORMULA_PREFIX).delete_suffix(ARRAY_FORMULA_SUFFIX) << '</f>'
        str << '<v>' << cell.formula_value.to_s << '</v>' unless cell.formula_value.nil?
      end

      # Serializes cells that are type inline_string
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def inline_string_serialization(cell, str = +'')
        str << 't="inlineStr"><is>'
        run_xml_string cell, str
        str << '</is>'
      end

      # Serializes cells that are type string
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def string(cell, str = +'')
        if cell.is_array_formula?
          array_formula_serialization cell, str
        elsif cell.is_formula?
          formula_serialization cell, str
        elsif !cell.ssti.nil?
          value_serialization 's', cell.ssti, str
        else
          inline_string_serialization cell, str
        end
      end

      # Serializes cells that are of the type richtext
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def richtext(cell, str)
        if cell.ssti.nil?
          inline_string_serialization cell, str
        else
          value_serialization 's', cell.ssti, str
        end
      end

      # Serializes cells that are of the type text
      # @param [Cell] cell The cell that is being serialized
      # @param [#<<] str A String, buffer or IO to append the serialization to.
      # @return [void]
      def text(cell, str)
        if cell.ssti.nil?
          inline_string_serialization cell, str
        else
          value_serialization 's', cell.ssti, str
        end
      end

      private

      def numeric(cell, str = +'')
        value_serialization 'n', cell.value, str
      end

      def value_serialization(serialization_type, serialization_value, str = +'')
        str << 't="' << serialization_type.to_s << '"' if serialization_type
        str << '><v>' << serialization_value.to_s << '</v>'
      end
    end
  end
end
