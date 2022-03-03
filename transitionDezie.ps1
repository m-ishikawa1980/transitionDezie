# pleasanter�o�^�p��Url���w�肵�܂��B
$url = 'http://localhost/api/items/99999/create'

# pleasanter�o�^�p��ApiKey���w�肵�܂��B
$ApiKey = 'xxxxxxxxxxxxxxxxxxxxxxx'

# �f�a�G����o�͂���xml�t�@�C����PATH���w�肵�܂��B
$source = 'C:\Users\�ΐ�w\Documents\��Ɨp\20220301_�f�W�G�ڍs�T���v��\db65.xml'

# �f�a�G������PATH���w�肵�܂��B�i�Y�t�t�@�C����o�^���邽�߁j
$deziePatch = 'C:\Users\�ΐ�w\Documents\��Ɨp\20220301_�f�W�G�ڍs�T���v��\dze'

# �f�a�G�̃t�B�[���hID��pleasanter�̍���ID�̑Ή����w�肵�܂��B
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

# �Y�t�t�@�C���̕ۑ���p�X�ҏW
function EditAttach-Path($libraryId, $fieldId, $recordId) {

    $path = $deziePatch + '\file\DB\' + $libraryId +'\' + $fieldId + '\' + $recordId
    return $path

}

# �Y�t�t�@�C����base64�ϊ�
function ConvertTo-Base64($filePath) {

    $file = Get-Item $filePath
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
    return $b64

}

# ���C������
function main {
    write-host 'main start'

    # xml�I�u�W�F�N�g�쐬
    $xmlObj = [System.Xml.XmlDocument](Get-Content $source)

    # ���C�u����ID���擾
    $libraryId = $xmlObj.dezie.library.GetAttribute('id')

    # record�P�ʂŃ��[�v
    foreach ($record in $xmlObj.dezie.library.'record-list'.record) {

        # ���R�[�hID���擾
        $recordId = $record.GetAttribute('id')    

        $Hash = @{}
        $AttachmentsHash = @{}
        $ClassHash = @{}
        $DateHash = @{}
        $CheckHash = @{}
        $NumHash = @{}
        $DescriptionHash = @{}

        Write-Host 'recordId = ' $recordId
        
        # xml��value�v�f����Pleasanter�֓o�^�������ҏW
        foreach($valueItem in $record.value){

            # �t�B�[���hID���擾
            $fieldId = $valueItem.GetAttribute('id')

            # �t�B�[���hID����Pleasanter�̍���ID�𔻒f���A�e���ڕ��ނ��ƂɕҏW
            switch -Wildcard ($ConvItemsId[$fieldId]) {

                'Attachments*' {
                    
                    # �Y�t�t�@�C���̓��C�u����ID�A�t�B�[���hID�A���R�[�hID���Ώۃt�@�C���𔻒f
                    $attachFile = EditAttach-Path $libraryId $fieldId $recordId                    

                    # base64�ɕϊ�
                    $b64 = ConvertTo-Base64 $attachFile
                    
                    # ���N�G�X�gjson�ҏW
                    $attachments = New-Object System.Collections.ArrayList
                    $attachments.Add(@{
                        Name = $valueItem.file.innerText
                        ContentType = 'application/octet-stream'
                        Base64 = $b64
                    })
                    $AttachmentsHash.Add($ConvItemsId[$fieldId],$attachments)

                }
                
                'Class*' {
                    
                    # ���N�G�X�gjson�ҏW
                    $ClassHash.Add($ConvItemsId[$fieldId],$valueItem.innerText)
                
                }

                'Num*' {
                    
                    # ���N�G�X�gjson�ҏW
                    $NumHash.Add($ConvItemsId[$fieldId],[int]$valueItem.innerText)
                
                }

                'Check*' {
                    
                    # ���N�G�X�gjson�p��bool�l�ɕϊ�
                    if($valueItem.innerText -eq '0'){

                        $CheckHash.Add($ConvItemsId[$fieldId],$false)

                    }else{

                        $CheckHash.Add($ConvItemsId[$fieldId],$true)

                    }
                }

                'Date*' {
                    
                    # ���N�G�X�gjson�p��datetime�ɕϊ�
                    $date = [DateTime]::ParseExact($valueItem.date,'yyyy-MM-dd', $null);
                    $DateHash.Add($ConvItemsId[$fieldId],$date)
                }

                'Description*' {
                    $DescriptionHash.Add($ConvItemsId[$fieldId],$valueItem.innerText)

                }
            }
        }


        # ���N�G�X�gjson�ҏW
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
        
        # API���N�G�X�g���s
        $res = Invoke-RestMethod -Uri $url -ContentType 'application/json' -Method POST -Body ${Byte}

        Write-Host 'Reg to Pleasanter Comp' 
    }

    write-host 'main end'
}

write-host 'start'

main

write-host 'end'