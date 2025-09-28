"use client"

import type React from "react"

import { useState } from "react"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Upload, FileSpreadsheet, CheckCircle, XCircle, Zap, Shield, BarChart3 } from "lucide-react"
import * as XLSX from "xlsx"

interface ProcessedData {
  [key: string]: any
  _date_group?: string
  _sort_key?: number
}

export default function ExcelDivisionSplitter() {
  const [coreFile, setCoreFile] = useState<File | null>(null)
  const [coCommFile, setCoCommFile] = useState<File | null>(null)
  const [coreStatus, setCoreStatus] = useState<string>("")
  const [coCommStatus, setCoCommStatus] = useState<string>("")
  const [corePreview, setCorePreview] = useState<any[]>([])
  const [coCommPreview, setCoCommPreview] = useState<any[]>([])

  // Function to extract numeric part for sorting
  const extractNumber = (roll: any): number => {
    if (!roll || roll === null || roll === undefined) return Number.POSITIVE_INFINITY
    const match = String(roll).match(/(\d+)/)
    return match ? Number.parseInt(match[1]) : Number.POSITIVE_INFINITY
  }

  // Function to check if a row is a date header
  const isDateRow = (row: any[]): boolean => {
    if (!row[0]) return false
    const firstCell = String(row[0]).trim()
    const datePattern =
      /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+\d{1,2}(st|nd|rd|th)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/
    return datePattern.test(firstCell)
  }

  // Function to process data while preserving date headers
  const processDataWithDates = (data: any[], rollCol: string): ProcessedData[] => {
    const processedRows: ProcessedData[] = []
    let currentDate: string | null = null

    data.forEach((row) => {
      const rowArray = Object.values(row)
      if (isDateRow(rowArray)) {
        currentDate = String(rowArray[0]).trim()
      } else {
        // Only process rows that have valid division and roll data
        if (row.Division && row[rollCol]) {
          const processedRow: ProcessedData = { ...row }
          processedRow._date_group = currentDate
          processedRow._sort_key = extractNumber(row[rollCol])
          processedRows.push(processedRow)
        }
      }
    })

    return processedRows
  }

  // Function to create division sheets with proper date grouping
  const createDivisionSheets = (data: ProcessedData[], rollCol: string): Blob => {
    const workbook = XLSX.utils.book_new()

    // Get unique divisions
    const divisions = [...new Set(data.filter((row) => row.Division).map((row) => row.Division))]

    divisions.forEach((division) => {
      const sheetData: any[] = []

      // Get all unique dates for this division
      const divisionRows = data.filter((row) => row.Division === division)
      const dates = [...new Set(divisionRows.map((row) => row._date_group).filter((date) => date))]

      // Get original column names (excluding helper columns)
      const originalColumns = Object.keys(data[0] || {}).filter((col) => !col.startsWith("_"))

      // Add header row
      sheetData.push(originalColumns)

      dates.forEach((date) => {
        if (date) {
          // Add date header row
          const dateRow = new Array(originalColumns.length).fill("")
          dateRow[0] = date
          sheetData.push(dateRow)

          // Add student records for this date, sorted by roll number
          const dateRecords = divisionRows
            .filter((row) => row._date_group === date)
            .sort((a, b) => (a._sort_key || 0) - (b._sort_key || 0))

          dateRecords.forEach((record) => {
            const cleanRecord = originalColumns.map((col) => record[col] || "")
            sheetData.push(cleanRecord)
          })

          // Add empty row after each date group
          sheetData.push(new Array(originalColumns.length).fill(""))
        }
      })

      // Remove last empty row if exists
      if (sheetData.length > 0 && sheetData[sheetData.length - 1].every((cell: any) => cell === "")) {
        sheetData.pop()
      }

      const worksheet = XLSX.utils.aoa_to_sheet(sheetData)
      const sheetName = String(division).substring(0, 31) // Excel sheet name limit
      XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
    })

    const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" })
    return new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
  }

  // Function to process uploaded file
  const processFile = async (file: File, type: "core" | "cocomm") => {
    try {
      const setStatus = type === "core" ? setCoreStatus : setCoCommStatus
      const setPreview = type === "core" ? setCorePreview : setCoCommPreview

      setStatus("Processing...")

      const arrayBuffer = await file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: "array" })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      // Convert to array of arrays to find header row
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][]

      let headerRow = -1
      for (let i = 0; i < Math.min(10, rawData.length); i++) {
        const row = rawData[i].map((cell) => String(cell || "").toLowerCase())
        if (row.includes("division") && (row.includes("roll") || row.includes("roll number"))) {
          headerRow = i
          break
        }
      }

      if (headerRow === -1) {
        setStatus('❌ Could not detect header row containing "Division" and "Roll"')
        return
      }

      // Extract column names from header row
      const columnNames = rawData[headerRow]

      // Create structured data
      const structuredData: any[] = []
      for (let i = 0; i < rawData.length; i++) {
        if (i === headerRow) continue // Skip header row itself

        const rowData: any = {}
        columnNames.forEach((colName, colIndex) => {
          rowData[colName] = rawData[i][colIndex] || ""
        })
        structuredData.push(rowData)
      }

      setPreview(structuredData.slice(0, 10))

      const rollCol = columnNames.includes("Roll Number") ? "Roll Number" : "Roll"

      if (columnNames.includes("Division") && (columnNames.includes("Roll") || columnNames.includes("Roll Number"))) {
        const processedData = processDataWithDates(structuredData, rollCol)
        const divisionsCount = new Set(processedData.filter((row) => row.Division).map((row) => row.Division)).size

        // Create download blob
        const blob = createDivisionSheets(processedData, rollCol)
        const url = URL.createObjectURL(blob)
        const a = document.createElement("a")
        a.href = url
        a.download = type === "core" ? "DJCSI Core Attendance.xlsx" : "DJCSI Co-Committee Attendance.xlsx"
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        URL.revokeObjectURL(url)

        setStatus(`✅ Excel processed! ${divisionsCount} sheets created with date headers preserved.`)
      } else {
        setStatus(`❌ Required columns not found. Available columns: ${columnNames.join(", ")}`)
      }
    } catch (error) {
      const setStatus = type === "core" ? setCoreStatus : setCoCommStatus
      setStatus(`❌ Error processing file: ${error instanceof Error ? error.message : "Unknown error"}`)
    }
  }

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>, type: "core" | "cocomm") => {
    const file = e.target.files?.[0]
    if (file) {
      if (type === "core") {
        setCoreFile(file)
        setCoreStatus("")
        setCorePreview([])
      } else {
        setCoCommFile(file)
        setCoCommStatus("")
        setCoCommPreview([])
      }
    }
  }

  return (
    <div className="min-h-screen gradient-bg">
      {/* Hero Section */}
      <div className="relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-br from-primary/10 via-transparent to-accent/10" />
        <div className="relative max-w-7xl mx-auto px-4 py-16">
          <div className="text-center space-y-8">
            <div className="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-primary/10 border border-primary/20 text-primary text-sm font-medium">
              <Zap className="h-4 w-4" />
              Excel Processing Tool
            </div>

            <div className="space-y-4">
              <h1 className="text-5xl md:text-7xl font-bold text-balance">
                Excel Division
                <span className="text-transparent bg-gradient-to-r from-primary to-accent bg-clip-text"> Splitter</span>
              </h1>
              <p className="text-xl text-muted-foreground max-w-2xl mx-auto text-balance">
                Transform your attendance data with intelligent division splitting while preserving date headers and
                maintaining perfect organization.
              </p>
            </div>

            <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <Shield className="h-4 w-4 text-success" />
                Secure Processing
              </div>
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <BarChart3 className="h-4 w-4 text-primary" />
                Smart Organization
              </div>
              <div className="flex items-center gap-2 text-sm text-muted-foreground">
                <FileSpreadsheet className="h-4 w-4 text-accent" />
                Excel Compatible
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="max-w-6xl mx-auto px-4 pb-16">
        <Tabs defaultValue="core" className="w-full">
          <div className="flex justify-center mb-8">
            <TabsList className="grid w-full max-w-md grid-cols-2 bg-card/50 backdrop-blur-sm border border-border/50">
              <TabsTrigger
                value="core"
                className="data-[state=active]:bg-primary data-[state=active]:text-primary-foreground"
              >
                Core Section
              </TabsTrigger>
              <TabsTrigger
                value="cocomm"
                className="data-[state=active]:bg-primary data-[state=active]:text-primary-foreground"
              >
                Co-Committee
              </TabsTrigger>
            </TabsList>
          </div>

          <TabsContent value="core" className="space-y-8">
            <div className="gradient-border">
              <Card className="border-0">
                <CardHeader className="text-center space-y-4">
                  <div className="mx-auto w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center">
                    <FileSpreadsheet className="h-8 w-8 text-primary" />
                  </div>
                  <div>
                    <CardTitle className="text-2xl">Core Section Processing</CardTitle>
                    <CardDescription className="text-base mt-2">
                      Upload your Excel file for Core attendance processing with intelligent division splitting
                    </CardDescription>
                  </div>
                </CardHeader>
                <CardContent className="space-y-6">
                  <div className="space-y-4">
                    <Label htmlFor="core-file" className="text-base font-medium">
                      Select Excel File (Core)
                    </Label>
                    <div className="flex flex-col sm:flex-row items-stretch gap-4">
                      <Input
                        id="core-file"
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => handleFileChange(e, "core")}
                        className="flex-1 h-12 bg-input/50 border-border/50 focus:border-primary"
                      />
                      <Button
                        onClick={() => coreFile && processFile(coreFile, "core")}
                        disabled={!coreFile}
                        className="h-12 px-8 bg-primary hover:bg-primary/90 text-primary-foreground font-medium"
                        size="lg"
                      >
                        <Upload className="h-5 w-5 mr-2" />
                        Process File
                      </Button>
                    </div>
                  </div>

                  {coreStatus && (
                    <Alert
                      className={`border-2 ${
                        coreStatus.includes("✅")
                          ? "border-success/50 bg-success/5"
                          : coreStatus.includes("❌")
                            ? "border-destructive/50 bg-destructive/5"
                            : "border-warning/50 bg-warning/5"
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        {coreStatus.includes("✅") && <CheckCircle className="h-5 w-5 text-success" />}
                        {coreStatus.includes("❌") && <XCircle className="h-5 w-5 text-destructive" />}
                        <AlertDescription className="text-base font-medium">{coreStatus}</AlertDescription>
                      </div>
                    </Alert>
                  )}

                  {corePreview.length > 0 && (
                    <div className="space-y-4">
                      <div className="flex items-center gap-2">
                        <BarChart3 className="h-5 w-5 text-primary" />
                        <Label className="text-base font-medium">Data Preview (First 10 rows)</Label>
                      </div>
                      <div className="rounded-xl border border-border/50 overflow-hidden bg-card/30 backdrop-blur-sm">
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead className="bg-muted/50">
                              <tr>
                                {Object.keys(corePreview[0] || {}).map((key) => (
                                  <th
                                    key={key}
                                    className="p-4 text-left border-r border-border/30 font-semibold text-foreground"
                                  >
                                    {key}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {corePreview.map((row, index) => (
                                <tr key={index} className="border-t border-border/30 hover:bg-muted/20">
                                  {Object.values(row).map((value, cellIndex) => (
                                    <td key={cellIndex} className="p-4 border-r border-border/30 text-muted-foreground">
                                      {String(value || "")}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>
            </div>
          </TabsContent>

          <TabsContent value="cocomm" className="space-y-8">
            <div className="gradient-border">
              <Card className="border-0">
                <CardHeader className="text-center space-y-4">
                  <div className="mx-auto w-16 h-16 rounded-2xl bg-accent/10 flex items-center justify-center">
                    <FileSpreadsheet className="h-8 w-8 text-accent" />
                  </div>
                  <div>
                    <CardTitle className="text-2xl">Co-Committee Processing</CardTitle>
                    <CardDescription className="text-base mt-2">
                      Upload your Excel file for Co-Committee attendance processing with intelligent division splitting
                    </CardDescription>
                  </div>
                </CardHeader>
                <CardContent className="space-y-6">
                  <div className="space-y-4">
                    <Label htmlFor="cocomm-file" className="text-base font-medium">
                      Select Excel File (Co-Committee)
                    </Label>
                    <div className="flex flex-col sm:flex-row items-stretch gap-4">
                      <Input
                        id="cocomm-file"
                        type="file"
                        accept=".xlsx,.xls"
                        onChange={(e) => handleFileChange(e, "cocomm")}
                        className="flex-1 h-12 bg-input/50 border-border/50 focus:border-accent"
                      />
                      <Button
                        onClick={() => coCommFile && processFile(coCommFile, "cocomm")}
                        disabled={!coCommFile}
                        className="h-12 px-8 bg-accent hover:bg-accent/90 text-accent-foreground font-medium"
                        size="lg"
                      >
                        <Upload className="h-5 w-5 mr-2" />
                        Process File
                      </Button>
                    </div>
                  </div>

                  {coCommStatus && (
                    <Alert
                      className={`border-2 ${
                        coCommStatus.includes("✅")
                          ? "border-success/50 bg-success/5"
                          : coCommStatus.includes("❌")
                            ? "border-destructive/50 bg-destructive/5"
                            : "border-warning/50 bg-warning/5"
                      }`}
                    >
                      <div className="flex items-center gap-3">
                        {coCommStatus.includes("✅") && <CheckCircle className="h-5 w-5 text-success" />}
                        {coCommStatus.includes("❌") && <XCircle className="h-5 w-5 text-destructive" />}
                        <AlertDescription className="text-base font-medium">{coCommStatus}</AlertDescription>
                      </div>
                    </Alert>
                  )}

                  {coCommPreview.length > 0 && (
                    <div className="space-y-4">
                      <div className="flex items-center gap-2">
                        <BarChart3 className="h-5 w-5 text-accent" />
                        <Label className="text-base font-medium">Data Preview (First 10 rows)</Label>
                      </div>
                      <div className="rounded-xl border border-border/50 overflow-hidden bg-card/30 backdrop-blur-sm">
                        <div className="overflow-x-auto">
                          <table className="w-full text-sm">
                            <thead className="bg-muted/50">
                              <tr>
                                {Object.keys(coCommPreview[0] || {}).map((key) => (
                                  <th
                                    key={key}
                                    className="p-4 text-left border-r border-border/30 font-semibold text-foreground"
                                  >
                                    {key}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {coCommPreview.map((row, index) => (
                                <tr key={index} className="border-t border-border/30 hover:bg-muted/20">
                                  {Object.values(row).map((value, cellIndex) => (
                                    <td key={cellIndex} className="p-4 border-r border-border/30 text-muted-foreground">
                                      {String(value || "")}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}
                </CardContent>
              </Card>
            </div>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  )
}
