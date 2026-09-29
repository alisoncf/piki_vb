Attribute VB_Name = "UTC_Date"
' $Archive: /APAIN/ACCUSAMP/Software/vbsampl 1.x/SysTime/UTC.bas $
'
' Created: 2004-07-15 by
'       Jens Cameron
'       Datainnovation i Lund AB
'       Scheelevägen 18
'       S-22363 Lund
'       Sweden
'       Voice:  +46 46 123839
'       FAX:    +46 46 2112795
'       E-mail jens.cameron(at)dilab.se
'
' Last modified:
'       $Modtime: 5/25/04 4:40p $
'       $Author: Jens $
'       $Revision: 4 $
'
' Modification history:
'       $Log: /APAIN/ACCUSAMP/Software/vbsampl 1.x/SysTime/UTC.bas $
'
' Purpose:
'       This module contains routines to retrieve UTC and local time from
'       the system and conversion between the two.
'
' Depends on:
'       None
'

Option Explicit

' System time structure
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

' Time zone information. Note that this one is defined wrong in API viewer.
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' Return current time as UTC
Public Function UTCTime() As Date
    Dim t As SYSTEMTIME
    
    GetSystemTime t
    UTCTime = DateSerial(t.wYear, t.wMonth, t.wDay) + TimeSerial(t.wHour, t.wMinute, t.wSecond) + t.wMilliseconds / 86400000#
End Function

' Return current time as local time
Public Function LocalTime() As Date
    Dim t As SYSTEMTIME
    
    GetLocalTime t
    LocalTime = DateSerial(t.wYear, t.wMonth, t.wDay) + TimeSerial(t.wHour, t.wMinute, t.wSecond) + t.wMilliseconds / 86400000
End Function

' Convert UTC time to local time for current time zone
Public Function UTCtoLocal(ByVal tDate As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim lRes As Long
    
    lRes = GetTimeZoneInformation(tzi)
    stUTC.wYear = Year(tDate)
    stUTC.wMonth = Month(tDate)
    stUTC.wDay = Day(tDate)
    stUTC.wHour = Hour(tDate)
    stUTC.wMinute = Minute(tDate)
    stUTC.wSecond = Second(tDate)
    stUTC.wMilliseconds = 0
    lRes = SystemTimeToTzSpecificLocalTime(tzi, stUTC, stLocal)
    UTCtoLocal = DateSerial(stLocal.wYear, stLocal.wMonth, stLocal.wDay) + TimeSerial(stLocal.wHour, stLocal.wMinute, stLocal.wSecond)
End Function

' Convert Local time to UTC time if possible.
' Note: The function may return 0, 1 or 2 datetime values.
' During switch from Standard to Daylight time, there is (usually) one hour missing
' and there is an invalid time range of one hour. This function returns an empty collection.
' When time changes from Daylight to Standard, there is an ambguity so one local time
' corresponds to two different UTC times. This function then returns a collection with
' two elements.
Public Function LocalToUTC(ByVal tDate As Date) As Collection
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lRes As Long
    Dim col As Collection
    Dim tUTC As Date
    
    Set col = New Collection
    lRes = GetTimeZoneInformation(tzi)
    tUTC = tDate + tzi.Bias / 1440
    If tzi.StandardDate.wMonth = 0 Then
        ' No daylight time -- no problem
        col.Add tUTC
    Else
        ' Assume we are fuzzing with +- one daylight bias, which is normally negative.
        ' So, datetimes will be ordered from earliest to latest in collection.
        If Round(UTCtoLocal(tUTC + tzi.DaylightBias / 1440) * 86400) = Round(tDate * 86400) Then
            col.Add CDate(tUTC + tzi.DaylightBias / 1440)
        End If
        If Round(UTCtoLocal(tUTC) * 86400) = Round(tDate * 86400) Then
            col.Add tUTC
        End If
        If Round(UTCtoLocal(tUTC - tzi.DaylightBias / 1440) * 86400) = Round(tDate * 86400) Then
            col.Add CDate(tUTC - tzi.DaylightBias / 1440)
        End If
    End If
    Set LocalToUTC = col
End Function
