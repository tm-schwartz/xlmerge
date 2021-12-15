module Merger
export main
using REPL
using REPL.TerminalMenus
using DataFrames
using DataFramesMeta
using Debugger
using XLSX

function readXlsm(path::String)::Dict
    xf = XLSX.readxlsx(path)
    sn = XLSX.sheetnames(xf)
    d = Dict{String,DataFrame}(nm => DataFrame(xf[nm][:], :auto) for nm in sn)
    return d
end

function getPaths()::NamedTuple
    println("Enter path to workbooks folder")
    pathToWorkbooks = readline()
    println("Enter number of header rows (if known)")
    nHeaderRows = readline()
    println("I have the Path to Workbooks Folder as $(pathToWorkbooks) and the number of header rows as $(nHeaderRows).\nIs this correct?[y/n]")
    resp = readline()
    if (resp == "n")
        println("Enter new Path to Workbooks (empty to keep same)")
        pathToWorkbooks = readline()
        println("Enter number of header rows (if known)")
        nHeaderRows = readline()
    end
    if (isdir(pathToWorkbooks) & length(nHeaderRows) >= 1)
        return (workbooks = pathToWorkbooks, rowskip = parse(Int64, nHeaderRows))
    elseif (isdir(pathToWorkbooks) & length(nHeaderRows) == 0)
        return (workbooks = pathToWorkbooks,)
    else
        error("$pathToWorkbooks may not exist")
    end
end

function selectMenu(options::Dict)::Dict
    sheetnames = sort(collect(keys(options)))
    while true
        menu = MultiSelectMenu(sheetnames, pagesize = 10, charset = :unicode)
        selectedindx = request("select sheets to merge. its recommended to choose sheets with the same layout:", menu)
        if (length(selectedindx) == 0)
            println("no sheets selected, merge canceled")
        end
        println("you selected:")
        for i in selectedindx
            println(" - ", sheetnames[i])
        end
        println("is this correct?[y/n]")
        resp = readline()
        if (resp == "n")
            continue
        end
        selected = [sheetnames[i] for i in selectedindx]
        for sheet in keys(options)
            if (sheet ∉ selected)
                delete!(options, sheet)
            end
        end
        return options
    end
end

function selectMenu(options::String)::Vector
    contents = readdir(options)
    while true
        menu = MultiSelectMenu(contents, pagesize = 10, charset = :unicode)
        selected = request("select files to merge:", menu)
        if (length(selected) == 0)
            println("no files selected, merge canceled")
        end
        println("you selected:")
        for i in selected
            println("- ", contents[i])
        end
        println("is this correct?[y/n]")
        resp = readline()
        if (resp == "n")
            continue
        end
        return [contents[i] for i in selected]
    end
end

function writeResult(sheets::Dict)
    println("Where should the result file be written to?(path to directory)")
    path = Nothing
    overwrite = false
    while true
        ptd = readline()
        if (isdir(ptd))
            path = joinpath(ptd, "AppendData.xlsx")
            if (isfile(path))
                println("$path exists, overwrite?[y/n]")
                overwrite = readline()
                overwrite == "y" ? overwrite = true : overwrite == "n" ? error("$path exists") : continue  # if overwrite == y then set = true else if overwrite == n then error, anything else loop again
            end
            break
        else
            println("$ptd was not found, please reenter")
        end
    end
    skeys = zip(map(x -> replace(x, " " => ""), collect(keys(sheets))), collect(keys(sheets)))
    sheetdat = (Symbol(sn) => (collect(eachcol(sheets[key])), names(sheets[key])) for (sn, key) in skeys)
    XLSX.writetable(path; overwrite = overwrite, sheetdat...)
end

function runmerge(path::String, rowskip::Int64)
    selectedpaths = selectMenu(path)
    sheets = readXlsm(joinpath(path, selectedpaths[1]))
    filteredtabs = selectMenu(sheets)
    for sheet in keys(filteredtabs)
        delete!(filteredtabs[sheet], rowskip+1:nrow(filteredtabs[sheet]))
    end
    for file in selectedpaths
        activewkbk = XLSX.readxlsx(joinpath(path, file))
        for sheet in keys(filteredtabs)
            toappend = DataFrame(activewkbk[sheet][:], :auto)
            filteredtabs[sheet] = vcat(filteredtabs[sheet], @linq toappend[rowskip+1:nrow(toappend), :] |> where(:x1 .!= Missing))
        end
    end
    writeResult(filteredtabs)
end

function runmerge(path::String)
    selectedpaths = selectMenu(path)
    sheets = readXlsm(joinpath(path, selectedpaths[1]))
    filteredtabs = selectMenu(sheets)
    headerrows = Dict{String,Int64}()
    for sheet in keys(filteredtabs)
        if get(headerrows, sheet, Nothing) == Nothing
            if (nrow(filteredtabs[sheet]) > 10)
                toprnt = filteredtabs[sheet][1:10, [:x1]]
            else
                toprnt = filteredtabs[sheet][1:nrow(filteredtabs[sheet]), [:x1]]
            end
            println("Here are the first $(nrow(toprnt)) rows of sheet $sheet:")
            println(toprnt)
            println("Please enter number of header rows:")
            resp = parse(Int64, readline())
            headerrows[sheet] = resp
        end
        rowskip = headerrows[sheet]
        delete!(filteredtabs[sheet], rowskip+1:nrow(filteredtabs[sheet]))
    end
    for file in selectedpaths
        activewkbk = XLSX.readxlsx(joinpath(path, file))
        for sheet in keys(filteredtabs)
            rowskip = headerrows[sheet]
            toappend = DataFrame(activewkbk[sheet][:], :auto)
            filteredtabs[sheet] = vcat(filteredtabs[sheet], @linq toappend[rowskip+1:nrow(toappend), :] |> where(:x1 .!= Missing))
        end
    end
    writeResult(filteredtabs)
end

function main()
    path::NamedTuple = getPaths()
    runmerge(path...)
end

end

