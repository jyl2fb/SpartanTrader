Module GlobalVariables
    Public ActiveDB As String = ""

    Public teamID As String = "19"
    Public currentDate As Date

    Public riskFreeRate As Double = 0
    Public maxMargin As Double = 0
    Public startDate As Date = "1/1/1"
    Public initialCAccount As Double = 0

    Public CAccount As Double = 0
    Public IPvalue As Double = 0
    Public APvalue As Double = 0
    Public lastPriceDownloadDate As Date = "1/1/1"
    Public margin As Double = 0
    Public TPVatStart As Double
    Public TaTPV As Double = 0

    Public CT As Transaction

    Public TPV As Double = 0
    Public lastTransactionDate As Date = "1/1/1"
    Public TE As Double = 0
    Public TEpercent As Double = 0
    Public sumTE As Double = 0
    Public lastTEUpDate As Date = "1/1/1"

    Public RecommendationFamily(11) As String
    Public MasterRecList As List(Of Transaction)
    Public IntermediaryRecList As List(Of Transaction)
    Public FinalRecList As List(Of Transaction)

    Public traderMode As String = ""

    Public secondsLeft As Integer
    Public tempNewDate As Date
    Public waitingForData As Boolean

End Module
