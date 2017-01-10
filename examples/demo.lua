----
--
-- A simple example of some of the features of the xlsxwriter.lua module.
--
-- Copyright 2014-2015, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

-- by hsq
--local workbook  = Workbook:new("demo.xlsx")
local workbook  = Workbook:new()

local worksheet = workbook:add_worksheet()

-- Widen the first column to make the text clearer.
worksheet:set_column("A:A", 20)

-- Add a bold format to use to highlight cells.
local bold = workbook:add_format({bold = true})

-- Write some simple text.
worksheet:write("A1", "Hello")

-- Text with formatting.
worksheet:write("A2", "World", bold)

-- Write some numbers, with row/column notation.
worksheet:write(2, 0, 123)
worksheet:write(3, 0, 123.456)

workbook:close()
-- by hsq
local s = workbook:get_zip_data()
print(s and #s or 'NIL')
if s then
    local f = assert(io.open('demo.xlsx', 'w'))
    assert(f:write(s))
    assert(f:close())
end
