# pleasanter登録用のUrlを指定します。
$url = 'http://localhost/api/items/99999/create'

# pleasanter登録用のApiKeyを指定します。
$ApiKey = 'xxxxxxxxxxxxxxxxxxxxxxx'

# デヂエから出力したxmlファイルのPATHを指定します。
$source = 'C:\Users\石川学\Documents\作業用\20220301_デジエ移行サンプル\db65.xml'

# デヂエ資源のPATHを指定します。（添付ファイルを登録するため）
$deziePatch = 'C:\Users\石川学\Documents\作業用\20220301_デジエ移行サンプル\dze'

# デヂエのフィールドIDとpleasanterの項目IDの対応を指定します。
$ConvItemsId = @{

    '4' = 'DateB'
    '5' = 'ClassA'
    '6' = 'ClassB'
    '117' = 'ClassC'
    '7' = 'ClassD'
    '8' = 'ClassF'
    '9' = 'DescriptionA'
    '10' = 'ClassE'
    '11' = 'DateA'
    '12' = 'CheckA'
    '136' = 'AttachmentsA'
    '142' = 'NumA'

}

# 添付ファイルの保存先パス編集
function EditAttach-Path($libraryId, $fieldId, $recordId) {

    $path = $deziePatch + '\file\DB\' + $libraryId +'\' + $fieldId + '\' + $recordId
    return $path

}

# 添付ファイルをbase64変換
function ConvertTo-Base64($filePath) {

    $file = Get-Item $filePath
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
    return $b64

}

# メイン処理
function main {
    write-host 'main start'

    # xmlオブジェクト作成
    $xmlObj = [System.Xml.XmlDocument](Get-Content $source)

    # ライブラリIDを取得
    $libraryId = $xmlObj.dezie.library.GetAttribute('id')

    # record単位でループ
    foreach ($record in $xmlObj.dezie.library.'record-list'.record) {

        # レコードIDを取得
        $recordId = $record.GetAttribute('id')    

        $Hash = @{}
        $AttachmentsHash = @{}
        $ClassHash = @{}
        $DateHash = @{}
        $CheckHash = @{}
        $NumHash = @{}
        $DescriptionHash = @{}

        Write-Host 'recordId = ' $recordId
        
        # xmlのvalue要素からPleasanterへ登録する情報を編集
        foreach($valueItem in $record.value){

            # フィールドIDを取得
            $fieldId = $valueItem.GetAttribute('id')

            # フィールドIDからPleasanterの項目IDを判断し、各項目分類ごとに編集
            switch -Wildcard ($ConvItemsId[$fieldId]) {

                'Attachments*' {
                    
                    # 添付ファイルはライブラリID、フィールドID、レコードIDより対象ファイルを判断
                    $attachFile = EditAttach-Path $libraryId $fieldId $recordId                    

                    # base64に変換
                    $b64 = ConvertTo-Base64 $attachFile
                    
                    # リクエストjson編集
                    $attachments = New-Object System.Collections.ArrayList
                    $attachments.Add(@{
                        Name = $valueItem.file.innerText
                        ContentType = 'application/octet-stream'
                        Base64 = $b64
                    })
                    $AttachmentsHash.Add($ConvItemsId[$fieldId],$attachments)

                }
                
                'Class*' {
                    
                    # リクエストjson編集
                    $ClassHash.Add($ConvItemsId[$fieldId],$valueItem.innerText)
                
                }

                'Num*' {
                    
                    # リクエストjson編集
                    $NumHash.Add($ConvItemsId[$fieldId],[int]$valueItem.innerText)
                
                }

                'Check*' {
                    
                    # リクエストjson用にbool値に変換
                    if($valueItem.innerText -eq '0'){

                        $CheckHash.Add($ConvItemsId[$fieldId],$false)

                    }else{

                        $CheckHash.Add($ConvItemsId[$fieldId],$true)

                    }
                }

                'Date*' {
                    
                    # リクエストjson用にdatetimeに変換
                    $date = [DateTime]::ParseExact($valueItem.date,'yyyy-MM-dd', $null);
                    $DateHash.Add($ConvItemsId[$fieldId],$date)
                }

                'Description*' {
                    $DescriptionHash.Add($ConvItemsId[$fieldId],$valueItem.innerText)

                }
            }
        }


        # リクエストjson編集
        $Hash.Add('ApiVersion',1.1)
        $Hash.Add('ApiKey',$ApiKey)
    
        if($AttachmentsHash.Count -gt 0) {
            $Hash.Add('AttachmentsHash',$AttachmentsHash)    
        }

        if($ClassHash.Count -gt 0) {
            $Hash.Add('ClassHash',$ClassHash)
        }

        if($NumHash.Count -gt 0) {
            $Hash.Add('NumHash',$NumHash)    
        }

        if($CheckHash.Count -gt 0) {
            $Hash.Add('CheckHash',$CheckHash)    
        }

        if($DateHash.Count -gt 0) {
            $Hash.Add('DateHash',$DateHash)    
        }

        if($DescriptionHash.Count -gt 0) {
            $Hash.Add('DescriptionHash',$DescriptionHash)    
        }

        $requestBody = $Hash | ConvertTo-Json -Depth 3

        #write-host $requestBody

        $Byte = [System.Text.Encoding]::UTF8.GetBytes($requestBody)
        
        # APIリクエスト発行
        $res = Invoke-RestMethod -Uri $url -ContentType 'application/json' -Method POST -Body ${Byte}

        Write-Host 'Reg to Pleasanter Comp' 
    }

    write-host 'main end'
}

write-host 'start'

main

write-host 'end'