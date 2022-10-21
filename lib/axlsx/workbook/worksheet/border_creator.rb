# encoding: UTF-8

module Axlsx
  class BorderCreator
    attr_reader :worksheet, :cells, :edges, :width, :color

    def initialize(worksheet, cells, args)
      @worksheet = worksheet
      @cells     = cells
      if args.is_a?(Hash)
        @edges = args[:edges] || Axlsx::Border::EDGES
        @width = args[:style] || :thin
        @color = args[:color] || '000000'
      else
        @edges = args || Axlsx::Border::Edges
        @width = :thin
        @color = '000000'
      end

      if @edges == :all
        @edges = Axlsx::Border::EDGES
      elsif @edges.is_a?(Array)
        @edges = (@edges.map(&:to_sym).uniq & Axlsx::Border::EDGES)
      else
        @edges = []
      end
    end

    def draw
      @edges.each do |edge| 
        worksheet.add_style(
          border_cells[edge], 
          {
            border: {style: @width, color: @color, edges: [edge]}
          }
        )
      end
    end

    private

    def border_cells
      {
        top:     "#{first_cell}:#{last_col}#{first_row}",
        right:   "#{last_col}#{first_row}:#{last_cell}",
        bottom:  "#{first_col}#{last_row}:#{last_cell}",
        left:    "#{first_cell}:#{first_col}#{last_row}",
      }
    end

    def first_cell
      @first_cell ||= cells.first.r
    end

    def last_cell
      @last_cell ||= cells.last.r
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
