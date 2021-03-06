----
--
-- A simple example of some of the features of the xlsxwriter.lua module.
--
-- Copyright 2014-2015, John McNamara, jmcnamara@cpan.org
--

local Workbook = require "xlsxwriter.workbook"

local autoSave = false
local filename = 'demo.xlsx'
local workbook  = Workbook:new(autoSave and filename --[[or '']])
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
local s = workbook:get_zip_data()
print(s and #s or 'NIL')
if s then
    local f = assert(io.open(filename, 'w'))
    assert(f:write(s))
    assert(f:close())
end
