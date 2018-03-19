//
//  iOSXLSCell.m
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/19.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import "iOSXLSCell.h"

#if ! __has_feature(objc_arc)
#error THIS CODE MUST BE COMPILED WITH ARC ENABLED!
#endif

@implementation iOSXLSCell

{
    char colString[3];
}
@synthesize row;
@synthesize type;
@dynamic colStr;
@synthesize col;
@synthesize str;
@synthesize val;

+ (iOSXLSCell *)blankCell
{
    return [iOSXLSCell new];
}

- (char *)colStr
{
    return colString;
}
- (void)setColStr:(char *)colS
{
    colString[0] = colS[0];
    colString[1] = colS[1];
    colString[2] = '\0';
}

- (void)show
{
    NSLog(@"%@", [self dump]);
}

- (NSString *)dump
{
    NSMutableString *s = [NSMutableString stringWithCapacity:128];
    
    const char *name;
    switch(type) {
        case cellBlank:        name = "cellBlank";        break;
        case cellString:    name = "cellString";    break;
        case cellInteger:    name = "cellInteger";    break;
        case cellFloat:        name = "cellFloat";        break;
        case cellBool:        name = "cellBool";        break;
        case cellError:        name = "cellError";        break;
        default:            name = "cellUnknown";    break;
    }
    
    [s appendString:@"====================\n"];
    [s appendFormat:@"CellType: %s row=%u col=%s/%u\n", name, row, colString, col];
    [s appendFormat:@"   string:    %@\n", str];
    
    switch(type) {
        case cellInteger:    [s appendFormat:@"     long:    %ld\n",    [val longValue]];    break;
        case cellFloat:        [s appendFormat:@"    float:    %lf\n",    [val doubleValue]];    break;
        case cellBool:        [s appendFormat:@"     bool:    %d\n",    [val boolValue]];    break;
        case cellError:        [s appendFormat:@"    error:    %ld\n",    [val longValue]];    break;
        default: break;
    }
    return s;
}

@end
