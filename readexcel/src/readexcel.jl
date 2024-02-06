module readexcel

using XLSX

function createfile(filename::String)
    XLSX.openxlsx(filename*".xlsx", mode="w") do xf
    end

end

function createfile(filename::String,name,data)
    XLSX.openxlsx(filename*".xlsx", mode="w") do xf
        sheet = xf[1]
        XLSX.rename!(sheet, name)
        sheet["B1"] = data[1]
        sheet["B2"] = parse(Int,data[2])
        sheet["A1"] = "Name"
        sheet["A2"] = "Age"
    
        # will add a row from "A5" to "E5"
        sheet["A5"] = collect(1:5) # equivalent to `sheet["A5", dim=2] = collect(1:4)`
  
        sheet["B3", dim=1] = collect(1:4)
    
        # will add a matrix from "A7" to "C9"
        sheet["A7:C9"] = [ 1 2 3 ; 4 5 6 ; 7 8 9 ]
    end

end

function readfile(filename::String)
    return XLSX.readxlsx("$(filename).xlsx")
end


XLSX.openxlsx("tony.xlsx", mode="rw") do xf
    sheet = xf[1]
    sheet["B1"] = "new data"
end

function fillexcelfile()
        # prompt to input 
    println("Name of your worksheet: ")  
    
    # Calling rdeadline() function 
    name = readline() 
    
    data = []
    println("Enter the your name: ") 
    buffer = readline()
    println(buffer)
    push!(data, buffer) 
    println("Enter your age: ") 
    push!(data, readline() )
    
    createfile("fakename",name,data)
end

end # module readexcel
