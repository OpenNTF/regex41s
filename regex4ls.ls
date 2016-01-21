Option Public
Option Declare

UseLSX "*javacon"

%REM Description
	This LotusScript library adds Support for Regular Expressions to LotusScript
	by supplying Wrappers around native Java classes.
	This lets you work with regular LotusScript objects without bothering with the
	nasty LS2J code.
%END REM

REM Version: v1.0.0

REM Author: Harald Albers, HS - Hamburger Software GmbH & Co. KG

%REM Copyright
	Copyright (C) 2010 HS - Hamburger Software GmbH & Co. KG. All rights reserved.

	This library is released under the Apache Licence Version 2.0, see
	http://www.apache.org/licenses/LICENSE-2.0.txt

	It is distributed in the hope that it will be useful, but without any warranty;
	without even the implied warranty of merchantibility or fitness for a particular purpose.
%END REM

%REM Examples
	Dim Matcher As New SimpleRegexMatcher

	' This example determines a match and prints "Borg. You will be assimilated"
	If Matcher.matches("We are the Borg. You will be assimilated.", "Borg.*assimilated") Then
		Print Matcher.match
	End If
	
	' This example prints true as case is ignored
	Matcher.IgnoreCase = True
	Print Matcher.matches("We are the BORG", "Borg")
%END REM

%REM
	A simple Regular Expression matcher that can determine whether a text matches
	a pattern. If it does, it gives you access to the matching part of the string
	via the	matches	property.
	
	The Regular Expressions used in this class follow java.util.regex.Pattern syntax
	as described in
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html

	To make the matching process case insensitive, either set the IgnoreCase property
	to true or use (?i) in your Regular Expression. 
%END REM
Class SimpleRegexMatcher

	Private jPatternClass As JavaClass
	Private jPatternFactoryMethod As JavaMethod
	Private jMatcher As JavaObject
	
	' Set this property to true if you want the matching process to be case insensitive.
	' You can also enable case insensitivity by including the sequence (?i) in your
	' Regular Expression.
	Public IgnoreCase As Boolean

	Private Property Get CASE_INSENSITIVE As Integer
		CASE_INSENSITIVE = getPatternConstant("CASE_INSENSITIVE")
	End Property
	
	Private Property Get UNICODE_CASE As Integer
		UNICODE_CASE = getPatternConstant("UNICODE_CASE")
	End Property
	
	' Returns a static constant from the java.util.regex.Pattern class by its name
	Private Function getPatternConstant(ConstantName As String) As Integer
		Dim jProperty As JavaProperty
		Set jProperty = jPatternClass.GetProperty(ConstantName)
		getPatternConstant = jProperty.GetValue()
	End Function
	
	Sub New()
		Dim jSession As New JavaSession
		Set jPatternClass = jSession.GetClass("java/util/regex/Pattern")
		Set jPatternFactoryMethod = jPatternClass.GetMethod("compile",_
			"(Ljava/lang/String;I)Ljava/util/regex/Pattern;")
	End Sub
	
	' Returns true if the Regular Expression Pattern is found in InputText
	Function matches(InputText As String, Pattern As String) As Boolean
		Set jMatcher = getMatcher(InputText, Pattern)
		matches = jMatcher.find()
	End Function
	
	' Returns a java.util.regex.Matcher object for the given InputText and Pattern.
	Private Function getMatcher(InputText As String, Pattern As String) As JavaObject
		Dim jPattern As JavaObject
		Set jPattern = jPatternFactoryMethod.Invoke(,Pattern, getPatternModifier())
		Set getMatcher = jPattern.matcher(InputText)
	End Function
	
	Private Function getPatternModifier() As Integer
		If IgnoreCase Then
			' the UNICODE_CASE flag extends case sensitivity support to non-ASCII characters
			getPatternModifier = CASE_INSENSITIVE Or UNICODE_CASE
		End If
	End Function
	
	' Returns the matching part of the last matching process.
	' If there was no match, an empty string is returned.
	Property Get match As String
		If jMatcher Is Nothing Then
			Error 1, TypeName(Me) & ".match is not available before calling match()."
		End If

		On Error Resume Next
		match = jMatcher.group()
	End Property
	
End Class