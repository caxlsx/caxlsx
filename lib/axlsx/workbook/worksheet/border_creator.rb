module Axlsx
  class BorderCreator
    def initialize(worksheet:, cells:, edges: nil, style: nil, color: nil)
      @worksheet = worksheet
      @cells = cells

      @edges = edges || :all
      @style = style || :thin
      @color = color || "000000"

      if @edges == :all
        @edges = Axlsx::Border::EDGES
      elsif !@edges.is_a?(Array)
        raise ArgumentError.new("Invalid edges provided, #{@edges}")
      else
        @edges = @edges.map { |x| x&.to_sym }.uniq

        if !(@edges - Axlsx::Border::EDGES).empty?
          raise ArgumentError.new("Invalid edges provided, #{edges}")
        end
      end
    end

    def draw
      if @cells.size == 1
        @worksheet.add_style(
          first_cell,
          {
            border: { style: @style, color: @color, edges: @edges }
          }
        )
      else
        @edges.each do |edge|
          @worksheet.add_style(
            border_cells[edge],
            {
              border: { style: @style, color: @color, edges: [edge] }
            }
          )
        end
      end
    end

    private

    def border_cells
      {
        top:     "#{first_cell}:#{last_col}#{first_row}",
        right:   "#{last_col}#{first_row}:#{last_cell}",
        bottom:  "#{first_col}#{last_row}:#{last_cell}",
        left:    "#{first_cell}:#{first_col}#{last_row}"
      }
    end

    def first_cell
      @first_cell ||= @cells.first.r
    end

    def last_cell
      @last_cell ||= @cells.last.r
    end

    def first_row
      @first_row ||= first_cell.scan(/\d+/).first
    end

    def first_col
      @first_col ||= first_cell.scan(/\D+/).first
    end

    def last_row
      @last_row ||= last_cell.scan(/\d+/).first
    end

    def last_col
      @last_col ||= last_cell.scan(/\D+/).first
    end
  end
end
