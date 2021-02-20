' Laura Waite CS 115 Winter 2020

Option Strict On
Option Explicit On

' Import statement, required for Form Closing event
Imports System.ComponentModel

Public Class frmMultipleChoice
    '-------------------------------------------------------------------------- 
    ' Environmental Board Game - Multiple Choice Form
    ' Laura Waite
    ' CS 115 Winter 2020 
    ' Date: 2020-03-20 
    ' Definition: A multiple choice question form for Environmental Board Game
    '--------------------------------------------------------------------------
    ' Declare public variables

    ' Integer value for multiple choice question answer
    Dim intAnswer As Integer

    ' Integer for random number for question generation
    Dim intRandom As Integer

    ' String for answer to multiple choice question
    Dim strAnswer As String

    ' Boolean to ensure submit button was clicked
    Dim blnClicked As Boolean

    Private Sub frmMultipleChoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '-------------------------------------------------------------------------- 
        ' Subroutine - frmMultipleChoice_Load
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: On-load event handler for Multiple Choice Form
        '--------------------------------------------------------------------------
        ' On form load, call QuestionCounter subroutine in MultiChoiceModule
        Call QuestionCounter()

        ' Determine which question to show
        If intCounter < 41 Then

            ' If counter is less than 41, call question from list in sequential order
            Call Questions(intCounter)

        Else
            ' If counter is 41 or higher, generate random number 1-40
            intRandom = CInt(Int(Rnd() * 40 + 1))

            ' Send random number to Questions subroutine for random question
            Call Questions(intRandom)

        End If

        ' First radio button is always checked
        rad1.Checked = True

        ' Focus on submit button
        btnSubmit.Focus()

        ' Set boolean - button has not been clicked
        blnClicked = False

        ' Arrange controls and resize form
        lblQuestion.Left = CInt((Me.ClientSize.Width / 2) - (lblQuestion.Width / 2))
        grpRadio.Left = CInt((Me.ClientSize.Width / 2) - (grpRadio.Width / 2))
        grpRadio.Top = CInt(lblQuestion.Bottom + 15)
        btnSubmit.Left = CInt((Me.ClientSize.Width / 2) - (btnSubmit.Width / 2))
        btnSubmit.Top = CInt(grpRadio.Bottom + 15)
        Me.Height = btnSubmit.Top + btnSubmit.Height + 60
        If lblQuestion.Width > grpRadio.Width Then
            Me.Width = lblQuestion.Left + lblQuestion.Width + 30
        Else
            Me.Width = grpRadio.Left + grpRadio.Width + 30
        End If


    End Sub

    Private Sub frmMultipleChoice_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        '-------------------------------------------------------------------------- 
        ' Subroutine - frmMultipleChoice_SizeChanged
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Size changed event handler for Multiple Choice Form
        '--------------------------------------------------------------------------
        ' Arrange controls and resize form
        lblQuestion.Left = CInt((Me.ClientSize.Width / 2) - (lblQuestion.Width / 2))
        grpRadio.Left = CInt((Me.ClientSize.Width / 2) - (grpRadio.Width / 2))
        grpRadio.Top = CInt(lblQuestion.Bottom + 15)
        btnSubmit.Left = CInt((Me.ClientSize.Width / 2) - (btnSubmit.Width / 2))
        btnSubmit.Top = CInt(grpRadio.Bottom + 15)
        Me.Height = btnSubmit.Top + btnSubmit.Height + 60
        If lblQuestion.Width > grpRadio.Width Then
            Me.Width = lblQuestion.Left + lblQuestion.Width + 30
        Else
            Me.Width = grpRadio.Left + grpRadio.Width + 30
        End If

    End Sub

    Sub Questions(intCounter As Integer)
        '-------------------------------------------------------------------------- 
        ' Subroutine - Questions
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Subroutine to determine which question is shown when a special 
        ' square is triggered in frmGameBoard
        '--------------------------------------------------------------------------
        ' Determine which question to show
        If intCounter = 1 Then

            ' Question 1
            lblQuestion.Text = "The greatest danger facing most endangered species is which of these?"

            ' Multiple choice options
            rad1.Text = "Hunting"
            rad2.Text = "Habitat loss"
            rad3.Text = "Disease"
            rad4.Text = " Capture for the pet trade"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Habitat Loss" & ControlChars.CrLf &
            ControlChars.CrLf & "Throughout the world, humans are burning, bulldozing, polluting and building on habitats of endangered species."

        ElseIf intCounter = 2 Then

            ' Question 2
            lblQuestion.Text = "Which of the following is not listed as an endangered species under the U.S. Endangered Species Act (note: not limited to American animals)?"

            ' Multiple choice options
            rad1.Text = "Bald eagle"
            rad2.Text = "Blue whale"
            rad3.Text = "Cheetah"
            rad4.Text = "Hawaiian monk seal"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Bald eagle" & ControlChars.CrLf &
            ControlChars.CrLf & "The bald eagle's population has recovered well, and the species was delisted in 1999."

        ElseIf intCounter = 3 Then

            ' Question 3
            lblQuestion.Text = "The Bronx Zoo once had a picture frame covered by cloth. It was labeled 'The most dangerous species in the world'. If you removed the cover, what did you see?"

            ' Multiple choice options
            rad1.Text = "A great white shark"
            rad2.Text = "A virus"
            rad3.Text = "A mirror"
            rad4.Text = "A mountain lion"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: A mirror" & ControlChars.CrLf &
            ControlChars.CrLf & "Showing the most dangerous species...humans."

        ElseIf intCounter = 4 Then

            ' Question 4
            lblQuestion.Text = "Which of the following substances is not biodegradable?"

            ' Multiple choice options
            rad1.Text = "Food waste"
            rad2.Text = "Styrofoam"
            rad3.Text = "Sawdust"
            rad4.Text = "Paper"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Styrofoam" & ControlChars.CrLf &
            ControlChars.CrLf & "Bacteria cannot break down styrofoam."

        ElseIf intCounter = 5 Then

            ' Question 5
            lblQuestion.Text = "What environmental problem is the result of chemical reactions in the atmosphere that involve sulfur?"

            ' Multiple choice options
            rad1.Text = "Smog"
            rad2.Text = "Holes in the ozone layer"
            rad3.Text = "Acid rain"
            rad4.Text = "Deforestation"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Acid rain" & ControlChars.CrLf &
            ControlChars.CrLf & "Acid rain can collect in bodies of fresh water, making them inhabitable to most life forms."

        ElseIf intCounter = 6 Then

            ' Question 6
            lblQuestion.Text = "What chemicals have been banned in most of the world because of their role in destroying the ozone layer?"

            ' Multiple choice options
            rad1.Text = "Metal oxides"
            rad2.Text = "DDT"
            rad3.Text = "Chlorofluorocarbons"
            rad4.Text = "Peroxides"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Chlorofluorocarbons" & ControlChars.CrLf &
            ControlChars.CrLf & "Chlorofluorocarbons, also known as CFCs, were once widely used in aerosol propellants."

        ElseIf intCounter = 7 Then

            ' Question 7
            lblQuestion.Text = "Who wrote 'Silent Spring', the first book documenting many of the problems caused by pesticides?"

            ' Multiple choice options
            rad1.Text = "Henry David Thoreau"
            rad2.Text = "Rachel Carson"
            rad3.Text = "Jaques Cousteau"
            rad4.Text = "Aldo Leopold"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Rachel Carson" & ControlChars.CrLf &
            ControlChars.CrLf & "Carson, a founder of the modern environmental movement, published works such as Silent Spring in the '60s."

        ElseIf intCounter = 8 Then

            ' Question 8
            lblQuestion.Text = "Which of the following is not a renewable resource?"

            ' Multiple choice options
            rad1.Text = "Gold"
            rad2.Text = "Fish"
            rad3.Text = "Timber"
            rad4.Text = "Wild mushrooms"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Gold" & ControlChars.CrLf &
            ControlChars.CrLf & "Renewable resources are naturally replenished over time, and gold is a finite resource which the earth cannot replenish."

        ElseIf intCounter = 9 Then

            ' Question 9
            lblQuestion.Text = "Which of the following practices produces the most organic water pollution?"

            ' Multiple choice options
            rad1.Text = "Humans bathing in the water"
            rad2.Text = "Recreational boating"
            rad3.Text = "Intensive livestock farming"
            rad4.Text = "Paper mills"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Intensive livestock farming" & ControlChars.CrLf &
            ControlChars.CrLf & "Millions of gallons of manure and other animal waste contaminate our waters yearly."

        ElseIf intCounter = 10 Then

            ' Question 10
            lblQuestion.Text = "'Ecology' comes from the Greek words 'oikos' and 'logos'. What is the literal translation of these words?"

            ' Multiple choice options
            rad1.Text = "Science of the world"
            rad2.Text = "Household science"
            rad3.Text = "Nature study"
            rad4.Text = "Natural life"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Household science" & ControlChars.CrLf &
            ControlChars.CrLf & "It is a word of fairly recent coinage. Logos means science and oikos means household, so ecology means the 'science of the household of nature'."

        ElseIf intCounter = 11 Then

            ' Question 11
            lblQuestion.Text = "What is light pollution?"

            ' Multiple choice options
            rad1.Text = "Light bulbs which were not properly disposed"
            rad2.Text = "Bright lights on motor vehicles"
            rad3.Text = "Outdoor lights that are left on all day, wasting electricity"
            rad4.Text = "Excessive artificial light in the night sky"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Excessive artificial light in the night sky" & ControlChars.CrLf &
            ControlChars.CrLf & "Light pollution can be seen in most developed cities. It's more than just an eyesore— light pollution has many negative affects on both humans and animals."

        ElseIf intCounter = 12 Then

            ' Question 12
            lblQuestion.Text = "Along with an excess of outdoor lighting fixtures, what causes the majority of light pollution?"

            ' Multiple choice options
            rad1.Text = "Lights that are not contained"
            rad2.Text = "Lights that are too big"
            rad3.Text = "Lights that are the wrong color"
            rad4.Text = "Lights that are too far from the ground"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Lights that are not contained" & ControlChars.CrLf &
            ControlChars.CrLf & "Light fixtures that do not have proper coverings, or are unshielded altogether, allow light to escape upwards when it should be focused down where it is actually used."

        ElseIf intCounter = 13 Then

            ' Question 13
            lblQuestion.Text = "Light pollution is what blocks people from viewing the starlit skies from cities."

            ' Multiple choice options
            rad1.Text = "True"
            rad2.Text = "False"
            rad3.Visible = False
            rad4.Visible = False

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: True" & ControlChars.CrLf &
            ControlChars.CrLf & "Light pollution makes it nearly impossible to see stars in major cities. The farther away you go from cities and towns, the brighter the stars appear."

        ElseIf intCounter = 14 Then

            ' Question 14
            lblQuestion.Text = "Many people use security lights to light their property through the night. What can people do to keep their houses safe and also to reduce light pollution?"

            ' Multiple choice options
            rad1.Text = "Use fewer lights by using lights with a wider range of light"
            rad2.Text = "Use dimmer lights"
            rad3.Text = "Use motion-activated lights"
            rad4.Text = "Use lights that are higher from the ground"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Use motion-activated lights" & ControlChars.CrLf &
            ControlChars.CrLf & "Security lights that turn on with a motion sensor deter trespassers while reducing light pollution and saving energy (and money.) Win-win!"

        ElseIf intCounter = 15 Then

            ' Question 15
            lblQuestion.Text = "The excess of nighttime light that comes with light pollution causes a disruption in our circadian rhythm. Which of the following best defines this term?"

            ' Multiple choice options
            rad1.Text = "Our body's 24-hour biological cycle"
            rad2.Text = "Our body's process of getting energy from sunlight"
            rad3.Text = "The way our circulatory system responds to light"
            rad4.Text = "Our eyes' period of relaxing after absorbing so much light during the day"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Our body's 24-hour biological cycle" & ControlChars.CrLf &
            ControlChars.CrLf & "When your circadian rhythm is disrupted, the immediate repercussions are that your sleeping/waking and digestive systems are thrown off; which can make you feel awful. Longterm effects include increased risk of obesity, cardiovascular events and even neurological issues."

        ElseIf intCounter = 16 Then

            ' Question 16
            lblQuestion.Text = "Some studies carried out in the early twenty-first century have shown that there is a correlation between light pollution and which of these diseases?"

            ' Multiple choice options
            rad1.Text = "Epilepsy"
            rad2.Text = "Breast cancer"
            rad3.Text = "Diabetes"
            rad4.Text = "Tuberculosis"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Breast cancer" & ControlChars.CrLf &
            ControlChars.CrLf & "There have been studies that show women who live in cities where it is possible to read outside at midnight are more likely to develop breast cancer than those who live where it is darker. Women who work the 'graveyard' shift and are under lights throughout the night also have a higher chance of having breast cancer."

        ElseIf intCounter = 17 Then

            ' Question 17
            lblQuestion.Text = "Light pollution can cause sleep disorders."

            ' Multiple choice options
            rad1.Text = "True"
            rad2.Text = "False"
            rad3.Visible = False
            rad4.Visible = False

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: True" & ControlChars.CrLf &
            ControlChars.CrLf & "Sleep disorders like insomnia and shift-work syndrome have been linked to light pollution. Insomnia is the inability to fall asleep. Shift-work syndrome (or shift-work sleep disorder) affects people who change shifts often or work at night. The conflict between their work hours and circadian rhythm can result in up to 4 hours less sleep per night than the average person."

        ElseIf intCounter = 18 Then

            ' Question 18
            lblQuestion.Text = "Animals also face many problems caused by light pollution. Which of the following is NOT one of them?"

            ' Multiple choice options
            rad1.Text = "Disrupted breeding times"
            rad2.Text = "Increased predation"
            rad3.Text = "Decreased amount of food"
            rad4.Text = "Disorientation"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Decreased amount of food" & ControlChars.CrLf &
            ControlChars.CrLf & "While food supply is not impacted, too much light can confuse animals, make it harder to hide from predators, or cause them to breed too early."

        ElseIf intCounter = 19 Then

            ' Question 19
            lblQuestion.Text = "Sea turtles are one type of animal that is largely affected by light pollution. It affects their ability to find which of these essentials?"

            ' Multiple choice options
            rad1.Text = "Food"
            rad2.Text = "Clean water"
            rad3.Text = "Mates"
            rad4.Text = "Nesting spots"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Nesting spots" & ControlChars.CrLf &
            ControlChars.CrLf & "Sea turtles are unable to lay their eggs if they do not find a beach that is dark enough. They can also be confused by lights and wander into cities instead of following the moonlight home."

        ElseIf intCounter = 20 Then

            ' Question 20
            lblQuestion.Text = "How does light pollution affect the behavior of migratory birds?"

            ' Multiple choice options
            rad1.Text = "They nest in cities instead of in the wild"
            rad2.Text = "They are drawn into cities where they fly into buildings"
            rad3.Text = "They take longer to reach their destination"
            rad4.Text = "They start their migration later"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: They are drawn into ciites where they fly into buildings" & ControlChars.CrLf &
            ControlChars.CrLf & "Birds are impacted by light pollution in many ways, including disrupted migration, breeding times, and injury."

        ElseIf intCounter = 21 Then

            ' Question 21
            lblQuestion.Text = "One of the greatest effects mankind ever had on the environment was this period of time in the Victorian Era, during which the United Kingdom caused unprecedented levels of air pollution. By what name was it known?"

            ' Multiple choice options
            rad1.Text = "The Age of Enlightenment"
            rad2.Text = "The Dark Ages"
            rad3.Text = "The Industrial Revolution"
            rad4.Text = "The Computer Age"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: The Industrial Revolution" & ControlChars.CrLf &
            ControlChars.CrLf & "While the Industrial Revolution was a very important catalyst for progress in many scientific fields during the 19th century, the long-term effects may have been one of the early steps leading to an environmental nightmare."

        ElseIf intCounter = 22 Then

            ' Question 22
            lblQuestion.Text = "Air pollution in the 19th and 20th century was heavily affected by fossil fuels. Which of these is an example of this type of fuel?"

            ' Multiple choice options
            rad1.Text = "Steam"
            rad2.Text = "Sunlight"
            rad3.Text = "Coal"
            rad4.Text = "Wind"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Coal" & ControlChars.CrLf &
            ControlChars.CrLf & "Fossil fuels such as coal, oil, and natural gas, were created over millions of years by the breakdown and build-up of materials beneath the earth's surface."

        ElseIf intCounter = 23 Then

            ' Question 23
            lblQuestion.Text = "Chlorofluorocarbons (also known as CFCs) were first created in the 1920s. These dangerous compounds are also known by what name?"

            ' Multiple choice options
            rad1.Text = "Freon"
            rad2.Text = "Neon"
            rad3.Text = "Radon"
            rad4.Text = "Plasma"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Freon" & ControlChars.CrLf &
            ControlChars.CrLf & "Freon was popularly used in refrigeration, aerosol cans, and even inhalers. This difficult-to-breakdown compound contributes to depletion of the ozone layer."

        ElseIf intCounter = 24 Then

            ' Question 24
            lblQuestion.Text = "Dichlorodiphenyltrichloroethane, or DDT, was created in the 1870s and later used as which of these, causing widespread damage to human health and natural ecosystems?"

            ' Multiple choice options
            rad1.Text = "Motor oil"
            rad2.Text = "Pesticide"
            rad3.Text = "Paint"
            rad4.Text = "Sweetener"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Pesticide" & ControlChars.CrLf &
            ControlChars.CrLf & "Outlined in Rachel Carson's 'Silent Spring', DDT was determined to be carcinogenic to humans and outright toxic to small animals, and although it was banned in the U.S. the chemical's effect on waterways, soil composition, and birdlife continued long after."

        ElseIf intCounter = 25 Then

            ' Question 25
            lblQuestion.Text = "The Bhopal Disaster of 1984, occurring at a Union Carbide plant in India, resulted in which toxicant being released over the city of Bhopal?"

            ' Multiple choice options
            rad1.Text = "Methyl isocyanate"
            rad2.Text = "Sulfur tetrafluoride"
            rad3.Text = "Hexaethyl tetraphosphate"
            rad4.Text = "Dichloroacetylene"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Methyl isocyanate" & ControlChars.CrLf &
            ControlChars.CrLf & "A Union Carbide pesticide plant in Bhopal, India emitted enough of this toxic gas in 1984 to suffocate thousands of local citizens, contaminate area water supplies and soil deposits, and decimate local river ecosystems. Thirty years after the incident, chemicals were still confirmed to be leaking from the plant."

        ElseIf intCounter = 26 Then

            ' Question 26
            lblQuestion.Text = "Which city's air pollution nearly cost a bid for the 2008 Summer Olympic Games?"

            ' Multiple choice options
            rad1.Text = "London, England"
            rad2.Text = "Athens, Greece"
            rad3.Text = "Beijing, China"
            rad4.Text = "Sydney, Australia"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Beijing, China" & ControlChars.CrLf &
            ControlChars.CrLf & "While China's insistence that air quality would improve in time for the twenty-ninth Olympiad was enough to grant the bid, samples taken during the event were more than double the expected amount."

        ElseIf intCounter = 27 Then

            ' Question 27
            lblQuestion.Text = "The explosion of the Deepwater Horizon and the subsequent oil spill were later determined to be the responsibility of which corporation?"

            ' Multiple choice options
            rad1.Text = "Exxon Mobil"
            rad2.Text = "Royal Dutch Shell"
            rad3.Text = "Chevron"
            rad4.Text = "BP"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: BP" & ControlChars.CrLf &
            ControlChars.CrLf & "The worst oil spill in history, resulting over two hundred million gallons of oil into the Gulf of Mexico, was found to be caused by negligence by BP— who were fined nearly nineteen billion dollars."

        ElseIf intCounter = 28 Then

            ' Question 28
            lblQuestion.Text = "Following an earthquake, Japan had an environmental catastrophe on their hands when which of these types of power plants was damaged?"

            ' Multiple choice options
            rad1.Text = "Hydroelectric"
            rad2.Text = "Coal"
            rad3.Text = "Solar"
            rad4.Text = "Nuclear"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Nuclear" & ControlChars.CrLf &
            ControlChars.CrLf & "The events in and around Fukushima Japan resulted in the largest nuclear catastrophe since Chernobyl, releasing significant amounts of nuclear material into the atmosphere for several days following the event."

        ElseIf intCounter = 29 Then

            ' Question 29
            lblQuestion.Text = "Citizens of Flint, Michigan found themselves unable to drink their water due to high concentrations of which of these in their piping?"

            ' Multiple choice options
            rad1.Text = "Streptococcus bacteria"
            rad2.Text = "Plutonium"
            rad3.Text = "Lead"
            rad4.Text = "Asbestos"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 3
            strAnswer = "Answer: Lead" & ControlChars.CrLf &
            ControlChars.CrLf & "This crisis began in 2014 and stretched across several years, starting with complaints about bad-tasting water that led investigators to discover an astounding amount of lead in the supply, as well as traces of E.Coli bacteria."

        ElseIf intCounter = 30 Then

            ' Question 30
            lblQuestion.Text = "During the Olympic Games in Rio de Janeiro, Brazil in 2016, concerns were raised about water quality in and outside the venues for the summer games. During this time, Olympic swimming pools turned which unhealthy shade?"

            ' Multiple choice options
            rad1.Text = "Red"
            rad2.Text = "Green"
            rad3.Text = "Yellow"
            rad4.Text = "Brown"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Green" & ControlChars.CrLf &
            ControlChars.CrLf & "While Brazil claimed to have one of the most environmentally-friendly games to date, their diving pools turned a sickly green prior to competitions, hinting at improper chemical usage and poor water conditions in indoor and outdoor venues alike."

        ElseIf intCounter = 31 Then

            ' Question 31
            lblQuestion.Text = "Which of the following is NOT one of the devastating effects associated with rising sea levels?"

            ' Multiple choice options
            rad1.Text = "Increased flooding from storm surges"
            rad2.Text = "Contamination of freshwater supplies"
            rad3.Text = "Disappearance of small island states"
            rad4.Text = "Easier access to harbors"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Easier access to harbors" & ControlChars.CrLf &
            ControlChars.CrLf & "While rising sea levels theoretically might increase the depth of some harbors, allowing easier access for bigger ships, this would not be a devastating effect."

        ElseIf intCounter = 32 Then

            ' Question 32
            lblQuestion.Text = "To reclaim land from the sea, the Dutch city of Amsterdam is constructing a new archipelago of artificial islands using a technique they have named for which delicious breakfast food?"

            ' Multiple choice options
            rad1.Text = "Porridge method"
            rad2.Text = "Pancake method"
            rad3.Text = "Bacon method"
            rad4.Text = "Grits method"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Pancake method" & ControlChars.CrLf &
            ControlChars.CrLf & "Using a 'pancake method', they have poured multiple layers of sand onto weak subsoil, allowing each to harden to sludge, and then covering them with the next 'pancake' until the islands reached two meters above the water level."

        ElseIf intCounter = 33 Then

            ' Question 33
            lblQuestion.Text = "Some coastal cities are investing in modern barrier systems to hold back water surges. Which of the following was NOT in use at the end of 2018?"

            ' Multiple choice options
            rad1.Text = "Giant floodgates with openings for ships to pass"
            rad2.Text = "Retractable floodgates that can be raised as needed"
            rad3.Text = "Modular inflatable flood barriers"
            rad4.Text = "Nuclear-driven giant windmills"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Nuclear-driven giant windmills" & ControlChars.CrLf &
            ControlChars.CrLf & "While giant nuclear-driven windmills are not yet in use, cities in different parts of the world are looking to a wide range of barrier technologies to hold back flooding from increasingly strong storm surges."

        ElseIf intCounter = 34 Then

            ' Question 34
            lblQuestion.Text = "In 2006, the Dutch created a new agency to manage periodic flooding of rivers in the Netherlands. What is it called?"

            ' Multiple choice options
            rad1.Text = "Room for the River"
            rad2.Text = "Dutch River Agency"
            rad3.Text = "Dutch River Management Corporation"
            rad4.Text = "River Level Management Agency"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 1
            strAnswer = "Answer: Room for the River" & ControlChars.CrLf &
            ControlChars.CrLf & "High barriers alone do not provide a comprehensive response to the water challenges of the future. Instead of focusing exclusively on keeping the water out, the Dutch are now taking a 'living with water' approach."

        ElseIf intCounter = 35 Then

            ' Question 35
            lblQuestion.Text = "One of the key problems New Orleans, Louisiana has been addressing in its post-Hurricane Katrina water plan is subsidence, which can be best described as which of the following?"

            ' Multiple choice options
            rad1.Text = "Flooding of low-lying areas"
            rad2.Text = "Settling or sudden sinking of land surface"
            rad3.Text = "Flooding of high-lying areas"
            rad4.Text = "Wasting of water assets"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Settling or sudden sinking of land surface" & ControlChars.CrLf &
            ControlChars.CrLf & "Subsidence occurs when groundwater is removed causing the soil to dry out and the resulting air pockets to compress and collapse."

        ElseIf intCounter = 36 Then

            ' Question 36
            lblQuestion.Text = "China's 'sponge cities' initiative is an attempt to increase the percentage of storm water runoff that is either captured, reused, or absorbed by the ground. Which is NOT one of the strategies being employed?"

            ' Multiple choice options
            rad1.Text = "Increased wetland areas"
            rad2.Text = "Artificial lakes"
            rad3.Text = "Rooftop gardens"
            rad4.Text = "Impermeable sidewalks"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 4
            strAnswer = "Answer: Impermeable sidewalks" & ControlChars.CrLf &
            ControlChars.CrLf & "China has started a 'sponge cities' initiative to help improve water absorption. Thirty pilot cities are implementing a range of strategies, such as replacing old, impermeable pavements with porous materials and creating artificial wetlands."

        ElseIf intCounter = 37 Then

            ' Question 37
            lblQuestion.Text = "Singapore has adopted a multi-approach plan to reduce the impact of rising sea levels. Which of the following is NOT part of that plan?"

            ' Multiple choice options
            rad1.Text = "Raising low-lying roads near coastal areas"
            rad2.Text = "Intensifying coastal development"
            rad3.Text = "Building a storage tank and diversion canal system"
            rad4.Text = "Increasing the desalination capacity of their water system"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Intensifying coastal development" & ControlChars.CrLf &
            ControlChars.CrLf & "Singapore has adopted many protective measures against rising sea-levels, from raising roads in low-lying areas and increasing the elevation of planned additions to their airport, to building a large storage tank and diversion canal complex."

        ElseIf intCounter = 38 Then

            ' Question 38
            lblQuestion.Text = "Rising sea levels present significant threats to agriculture in many coastal areas. Which of the following is one of the principal factors?"

            ' Multiple choice options
            rad1.Text = "Decreasing salinity of the soil"
            rad2.Text = "Increasing salinity of the soil"
            rad3.Text = "Increasing levels of calcium in the soil"
            rad4.Text = "Decreasing levels of borax in the soil"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: Increasing salinity of the soil" & ControlChars.CrLf &
            ControlChars.CrLf & "Rising sea levels and flooding impact agriculture in several ways, from the loss of growing crops to the erosion of arable land and increased salinity of the soil."

        ElseIf intCounter = 39 Then

            ' Question 39
            lblQuestion.Text = "A 2018 study of the impact of rising sea levels on the Marshall Islands in the Pacific Ocean predicted that they could become uninhabitable by as early as which year?"

            ' Multiple choice options
            rad1.Text = "2050"
            rad2.Text = "2030"
            rad3.Text = "2100"
            rad4.Text = "2075"

            ' Set visibility for 3 & 4 to prevent missing choices when last question had 2 choices only
            rad3.Visible = True
            rad4.Visible = True

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: 2030" & ControlChars.CrLf &
            ControlChars.CrLf & "Low-lying island states, such as the Marshall Islands would be among the hardest hit by rising sea levels. A 2018 study conducted by the US military predicts that the islands could be uninhabitable as early as 2030. "

        ElseIf intCounter = 40 Then

            ' Question 40
            lblQuestion.Text = "If carbon emissions were stopped completely tomorrow, would that immediately halt the rise of sea levels around the world?"

            ' Multiple choice options
            rad1.Text = "Yes"
            rad2.Text = "No"
            rad3.Visible = False
            rad4.Visible = False

            ' Answer
            intAnswer = 2
            strAnswer = "Answer: No" & ControlChars.CrLf &
            ControlChars.CrLf & "Recent studies suggest that reaction time for sea level rise is much slower than for global temperatures. A 2018 paper published by climate scientists at Oregon State University predicts that the sea levels will continue to rise for up to a thousand years."

        End If
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        '-------------------------------------------------------------------------- 
        ' Subroutine - btnSubmit_Click
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: Click event handler for Submit button
        '--------------------------------------------------------------------------
        ' Set boolean - button has been clicked
        blnClicked = True

        ' Variable for selected radio button value
        Dim intGuess As Integer

        ' Set radio button values
        If rad1.Checked Then
            intGuess = 1
        ElseIf rad2.Checked Then
            intGuess = 2
        ElseIf rad3.Checked Then
            intGuess = 3
        Else
            intGuess = 4
        End If

        ' Set rules for correct or incorrect answers
        If intGuess = intAnswer Then

            ' Answer is correct
            blnAnswer = True

            ' Call QuestionResult subroutine in MultiChoiceModule
            Call QuestionResult()

            ' Display message box with answer information
            MsgBox(strAnswer,, "Correct!")

        Else
            ' Answer is incorrect
            blnAnswer = False

            ' Call QuestionResult subroutine in MultiChoiceModule
            Call QuestionResult()

            ' Display message box with answer information
            MsgBox(strAnswer,, "Sorry!")

        End If

        ' After user clicks message box "OK" button, close form
        Me.Close()

    End Sub

    Private Sub frmMultipleChoice_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        '-------------------------------------------------------------------------- 
        ' Subroutine - frmMultipleChoice_Closing
        ' Laura Waite
        ' CS 115 Winter 2020 
        ' Date: 2020-03-20 
        ' Definition: On close event handler for Multiple Choice Form
        '--------------------------------------------------------------------------
        ' If form is closed without submitting an answer to the question
        If blnClicked = False Then

            If blnP1Turn Then
                ' Player 1 loses their next turn
                blnP1LoseATurn = True
            Else
                ' Player 2 loses their next turn
                blnP2LoseATurn = True
            End If

            ' Display message encouraging game play
            MsgBox("You didn't submit an answer, so you've lost a turn.",, "You gotta play to win!")

        Else
            ' Player answered question, game play continues normally
            blnP1LoseATurn = False
            blnP2LoseATurn = False

        End If
    End Sub

End Class