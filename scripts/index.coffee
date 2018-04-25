module.exports = (robot) ->
    
    robot.respond /ping$/i, (msg) ->
        msg.send "ping"
        fpath = '.\\xls' + '\\dbQuery.xlsx'
        qData=xsltojs fpath
    
        #console.log "hi"
        console.log qData

xsltojs = (path) ->
        #console.log 'In ConfigFile Funciton'
        #fpath = '.\\xls' + '\\configData.xlsx'
        #console.log path
        XLSX  = require ('xlsx') 
        workbook = XLSX.readFile(path)
        sheet_name_list = workbook.SheetNames
        data = []
        
        for k,v in sheet_name_list
            #console.log k + ' ' + v
            worksheet =  workbook.Sheets[k]       
            #console.log(XLSX.utils.sheet_to_json(worksheet,{raw:true}))     
            data = XLSX.utils.sheet_to_json(worksheet,{raw:true})                       
        return data

makeQuery = (data, changeStr) ->
        console.log 'old : ' + changeStr
        cnt = 0
        
        while cnt < data.length
            if cnt > 0 
                refstr= '{{ref_text' + cnt + '}}'
            else
                refstr= '{{ref_text}}'
            # console.log refstr
            # console.log data[cnt]
            changeStr = changeStr.replace refstr , data[cnt]
            cnt++
        console.log 'new : ' + changeStr
        return changeStr