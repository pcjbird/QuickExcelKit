Pod::Spec.new do |s|
    s.name             = "QuickExcelKit"
    s.version          = "1.0.0"
    s.summary          = "A great & strong microsoft excel parser for iOS, support csv/xls/xlsx format. 一款微软 Excel 文件解析器，支持 csv/xls/xlsx 格式。"
    s.description      = <<-DESC
    A great & strong microsoft excel parser for iOS, support csv/xls/xlsx format. 一款 iOS 平台上的微软 Excel 文件解析器，支持 csv/xls/xlsx 格式, 该项目支持 Cocoapods 的 subspec。
    DESC
    s.homepage         = "https://github.com/pcjbird/QuickExcelKit"
    s.license          = 'MIT'
    s.author           = {"pcjbird" => "pcjbird@hotmail.com"}
    s.source           = {:git => "https://github.com/pcjbird/QuickExcelKit.git", :tag => s.version.to_s, :submodules => true}
    s.social_media_url = 'http://www.lessney.com'
    s.requires_arc     = true
    s.documentation_url = 'https://github.com/pcjbird/QuickExcelKit/blob/master/README.md'
    s.screenshot       = 'https://github.com/pcjbird/QuickExcelKit/blob/master/logo.png'

    s.platform         = :ios, '8.0'
    s.frameworks       = 'Foundation', 'UIKit'
#s.preserve_paths   = ''
    s.source_files     = 'QuickExcelKit/QuickExcelKit.h'
    s.prefix_header_file = 'QuickExcelKit/Supporting Files/QuickExcelKit-Prefix.pch'

    s.default_subspec = 'QuickExcelReaderUtil'

    s.pod_target_xcconfig = { 'OTHER_LDFLAGS' => '-lObjC'}

    s.subspec 'ZXLSXParser' do |ss|
       ss.source_files = 'ZXLSXReader/ZXLSXReader/ZXLSXParser/xmlParser/*.{h,m}', 'ZXLSXReader/ZXLSXReader/ZXLSXParser/xmlParser/models/*.{h,m}'
       ss.public_header_files = 'ZXLSXReader/ZXLSXReader/ZXLSXParser/xmlParser/models/ZContent.h','ZXLSXReader/ZXLSXReader/ZXLSXParser/xmlParser/ZXLSXParser.h'
       ss.dependency 'SSZipArchive'
    end

    s.subspec 'iOSlibxls' do |ss|
        ss.source_files = 'iOSlibxls/*.{h,m}', 'iOSlibxls/libxls/include/libxls/*.{h}', 'iOSlibxls/src/*.{c}'
        ss.public_header_files = 'iOSlibxls/iOSXLSReader.h','iOSlibxls/iOSXLSCell.h'
        ss.header_mappings_dir = 'iOSlibxls'
        ss.pod_target_xcconfig = {
            'HEADER_SEARCH_PATHS' => '$(inherited) ./iOSlibxls ./iOSlibxls/libxls/incude'
        }
    end

    s.subspec 'QuickExcelReaderUtil' do |ss|
        ss.source_files = 'QuickExcelKit/*.{h,m}', 'QuickExcelKit/CSVParser/*.{h,m}', 'QuickExcelKit/Supporting Files/*.{pch}'
        ss.public_header_files = 'QuickExcelKit/QuickExcelReaderUtil.h'
        ss.dependency 'QuickExcelKit/ZXLSXParser'
        ss.dependency 'QuickExcelKit/iOSlibxls'
    end

end
