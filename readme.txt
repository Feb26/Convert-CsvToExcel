���@�\
�ECSV�t�@�C�����AExcel�t�@�C���ɕϊ����܂�
�E���ׂẴf�[�^�̓e�L�X�g�Ƃ��ēǂݍ��܂�܂��i���t�␔�l�Ƃ��Ă͓ǂݍ��܂�܂���j
�E�V�[�g���ɂ͌��̃t�@�C�������ݒ肳��܂�
	�E���̃t�@�C�����ɃV�[�g���ɐݒ�ł��Ȃ��������܂܂��A�܂���32�����𒴉߂��Ă���ꍇ�A�f�t�H���g�l(Sheet1)�̂܂܂ƂȂ�܂�
�E�e���̓t�@�C���̕����G���R�[�f�B���O�Ɖ��s�R�[�h�͎����I�ɔ��ʂ���܂�



���g�p���@
�ȉ���A�AB�����ꂩ�̕��@�ŕϊ��ł��܂��B

A. �h���b�O�A���h�h���b�v(D&D)�����t�@�C����ϊ�����
	1. run.bat�Ƀt�@�C����f�B���N�g����I������D&D����
		�E�����t�@�C���𓯎���D&D�ł��܂�
	2. D&D�����̂��f�B���N�g���̏ꍇ�A�\�����ꂽ�r���[���g�p���āA�ϊ�����t�@�C����I������
		�E�t�@�C���̏ꍇ�A�����I�ɕϊ����Aoutput�f�B���N�g���ɏo�͂���܂�
	3. OK�{�^������������
		�Eoutput�f�B���N�g���ɏo�͂���܂�

B. input�f�B���N�g���z���̃t�@�C����ϊ�����
	1. input�f�B���N�g���ɕϊ��������t�@�C����t�@�C�����܂ރf�B���N�g�����R�s�[����
	2. run.bat���_�u���N���b�N����
	3. �\�����ꂽ�r���[���g�p���āA�ϊ�����t�@�C����I������
	4. OK�{�^������������
		�Einput�f�B���N�g���̍\����ۂ����܂�output�f�B���N�g���ɏo�͂���܂�



���ݒ�
�E�ݒ�l��settings.json�ŕύX�\�ł�
�E�����l�͈ȉ��̂悤�ɂȂ��Ă��܂�

{
	"inputFile": {
		"extension": "*",
		"delimiter": ",",
		"textQualifier": "\""
	},

	"outputFile": {
		"font": "�l�r �o�S�V�b�N",
		"adjustColumnWidth": true
	},

	"path": {
		"inputDir": ".\\input",
		"outputDir": ".\\output"
	}
}

| ����              | ����                                                 | ��                                       |
| :---------------- | :--------------------------------------------------- | :--------------------------------------- |
| extension         | ���̓t�@�C���̊g���q���w�肵�܂��B                   | "*"  "csv"  "tsv"  "txt"                 |
| delimiter         | ���̓t�@�C���̋�؂蕶�����w�肵�܂��B               | ","  ";"  "\t"                           |
| textQualifier     | ���̓t�@�C���̕�����̈͂ݕ������w�肵�܂��B         | "\""  "'"                                |
| font              | �o�̓t�@�C���̃u�b�N�̃t�H���g���w�肵�܂��B��1      | "�l�r�@�o�S�V�b�N"  "Tahoma"  "Consolas" |
| adjustColumnWidth | �o�̓t�@�C���̗񕝂𒲐߂���ꍇtrue���w�肵�܂��B   | true  false                              |
| inputDir          | ���̓t�@�C���̃��[�g�f�B���N�g���̃p�X���w�肵�܂��B | "C:\\test\\input"                        |
| outputDir         | �o�̓t�@�C���̃��[�g�f�B���N�g���̃p�X���w�肵�܂��B | "C:\\test\\output"                       |

��1 �t�H���g�ɂ���
	�E���s���ɃC���X�g�[������Ă���t�H���g�����w�肵�Ă�������
	�EPowerShell�ɂĉ��L�R�}���h�����s����ƁA�t�H���g���̈ꗗ���m�F���邱�Ƃ��ł��܂�
	�E[void][reflection.assembly]::LoadWithPartialName("System.Drawing");[System.Drawing.FontFamily]::Families



���f�B���N�g���\��
./
|-- 
|-- input                                    �g�p���@B�œ��̓t�@�C����u���f�B���N�g��
|-- lib                                      �g�p���Ă��郉�C�u������dll�Ȃǂ��i�[����f�B���N�g��
|-- output                                   �o�̓t�@�C����Excel���o�͂���f�B���N�g��
|-- src                                      �\�[�X�R�[�h���i�[����f�B���N�g��
|   |-- Convert-AllCsvInInputDir2Excel.ps1   �g�p���@B�̃\�[�X�R�[�h
|   |-- Convert-AllCsvOfArgs2Excel.ps1       �g�p���@A�̃\�[�X�R�[�h
|   |-- Convert-Csv2Excel.ps1                CSV��Excel�ɕϊ����鏈���̃R�[�h
|   `-- Resolve-Encoding.ps1                 �����G���R�[�f�B���O�̎������菈���̃R�[�h
|-- log.txt                                  ���O�t�@�C��
|-- readme.txt                               �{�t�@�C��
|-- run.bat                                  �N���Ɏg�p����t�@�C��
`-- settings.json                            �ݒ�t�@�C��



������m�F��
| Software                 | Ver.           |
| :----------------------- | :------------- |
| Windows 10 Enterprise    | 1709           |
| PowerShell               | 5.1.16299.1004 |
| .NET Framework           | 4.5            |



���g�p���C�u����
| Name        | Ver.       | Path                                                                   |
| :---------- | :--------- | :--------------------------------------------------------------------- |
| EPPlus      | 4.5.3.1    | https://www.nuget.org/packages/EPPlus/                                 |



���ύX����
| Ver.  | �ύX���e                                                                                  |
| :---- | :---------------------------------------------------------------------------------------- |
| 1.0   | �V�K�쐬                                                                                  |
| 1.1   | Com�I�u�W�F�N�g�̉���R��ɂ��v���Z�X���I�����̉��P                                   |
| 1.2   | D&D�ɂ��ϊ��@�\�̒ǉ�                                                                   |
| 1.3   | UTF-8�Ή�                                                                                 |
| 1.4   | UNC�Ή�                                                                                   |
| 1.5   | Excel�̃v���Z�X���c���Ă��܂����ꍇ�ɂ�Kill������@�̉��P                                 |
| 1.6   | COM�I�u�W�F�N�g���������ǉ����v���Z�XKill�������폜                                     |
| 1.7   | �����G���R�[�f�B���O�̎�������@�\��ǉ�                                                  |
| 1.8   | �f�B���N�g���\���̕ύX�AD&D���̑��x���P                                                   |
| 2.0   | Microsoft.Office.Interop.Excel����EPPlus�ɕύX�A�ݒ荀�ڂ̒ǉ��i���s�R�[�h�Ȃǁj          |
| 2.1   | ���s�R�[�h�̎�������@�\��ǉ�                                                            |
| 2.2   | ���O�o�͋@�\�̒ǉ�                                                                        |
| 2.2.1 | �V�[�g���̐ݒ�Ɏ��s�����ۂ̃n���h�����O�̃f�O���C��                                      |