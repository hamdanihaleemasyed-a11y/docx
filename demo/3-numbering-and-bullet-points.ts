const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  PageNumber, Footer, Header, BorderStyle, ShadingType
} = require('docx');
const fs = require('fs');

// Helper functions
const heading1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 480, after: 200 },
  children: [new TextRun({ text, bold: true, size: 32, font: "Times New Roman", color: "1F3864" })]
});

const heading2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 300, after: 160 },
  children: [new TextRun({ text, bold: true, size: 26, font: "Times New Roman", color: "2E5B9A" })]
});

const thesis = (text) => new Paragraph({
  spacing: { before: 160, after: 160 },
  border: {
    left: { style: BorderStyle.THICK, size: 12, color: "2E5B9A", space: 8 }
  },
  indent: { left: 360 },
  children: [
    new TextRun({ text: "Thesis: ", bold: true, italics: true, size: 24, font: "Times New Roman", color: "2E5B9A" }),
    new TextRun({ text, italics: true, size: 24, font: "Times New Roman" })
  ]
});

const para = (children) => new Paragraph({
  spacing: { before: 120, after: 120, line: 360 },
  alignment: AlignmentType.JUSTIFIED,
  children
});

const t = (text, opts = {}) => new TextRun({ text, size: 24, font: "Times New Roman", ...opts });
const bold = (text) => t(text, { bold: true });
const italic = (text) => t(text, { italics: true });

const cite = (quote, act, scene, line) => [
  t('"'),
  italic(quote),
  t(`" (${act}, ${scene}, ${line})`)
];

const pageBreak = () => new Paragraph({
  pageBreakBefore: true,
  children: [new TextRun("")]
});

const titlePage = () => [
  new Paragraph({ spacing: { before: 1440, after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Hamlet by William Shakespeare", size: 48, bold: true, font: "Times New Roman", color: "1F3864" })] }),
  new Paragraph({ spacing: { before: 0, after: 240 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "A Comprehensive Thematic Analysis", size: 32, font: "Times New Roman", italics: true, color: "2E5B9A" })] }),
  new Paragraph({ spacing: { before: 480, after: 120 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Eight Major Themes with Textual References", size: 26, font: "Times New Roman", color: "444444" })] }),
  new Paragraph({ spacing: { before: 80, after: 80 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Prepared for Examination (20 Marks)", size: 24, font: "Times New Roman", color: "666666" })] }),
];

// ============================================================
// THEME 1: Tragic Elements (Greek, Roman, English)
// ============================================================
const theme1 = () => [
  pageBreak(),
  heading1("Theme 1: Hamlet — The Tragic Elements (Greek, Roman and English)"),
  thesis("Hamlet synthesises the Greek principle of hamartia, the Roman tradition of Senecan revenge, and the English Elizabethan fascination with political corruption into a tragedy that transcends any single cultural tradition, establishing it as the definitive tragic drama of the Western canon."),

  heading2("The Greek Tragic Elements"),
  para([t("Shakespeare's "), italic("Hamlet"), t(" draws profoundly upon the Aristotelian model of tragedy as defined in the "), italic("Poetics"), t(". Aristotle prescribed that a tragic hero must be a man of noble stature who falls from greatness due to a fatal flaw — "), italic("hamartia"), t(" — and whose downfall evokes pity and fear in the audience, ultimately producing "), italic("catharsis"), t(". Hamlet satisfies all these criteria with remarkable precision. He is a prince of Denmark, eloquent, educated at Wittenberg, and morally sensitive — qualities that mark him as a man set above common humanity. His hamartia is his tendency toward over-reflection, his paralysing capacity for philosophical speculation that prevents decisive action. Shakespeare himself signals this when Hamlet confesses, "), ...cite("I do not know / Why yet I live to say 'This thing's to do,' / Sith I have cause, and will, and strength, and means / To do't", "Act IV", "Scene 4", "lines 43–46"), t(". This self-awareness in the face of inaction is classically Greek — the hero sees his flaw but cannot overcome it.")]),
  para([t("The Greek concept of "), italic("anagnorisis"), t(" — the moment of tragic recognition — appears in Hamlet's graveyard meditation, where he confronts the skull of Yorick and recognises the absolute democracy of death: "), ...cite("Imperious Caesar, dead and turn'd to clay, / Might stop a hole to keep the wind away", "Act V", "Scene 1", "lines 213–214"), t(". This recognition triggers a philosophical calm that brings Hamlet closer to the acceptance necessary for the play's resolution. Similarly, "), italic("peripeteia"), t(", the reversal of fortune, occurs structurally when Hamlet's plans constantly produce the opposite of their intended effect — his feigned madness alienates Ophelia, his delay allows Claudius to consolidate power, and his accidental killing of Polonius sets in motion the revenge chain that ultimately destroys him.")]),

  heading2("Roman (Senecan) Tragic Elements"),
  para([t("The Roman tradition — specifically the influence of Lucius Annaeus Seneca — shaped Elizabethan tragedy decisively, and "), italic("Hamlet"), t(" is saturated with Senecan conventions. Seneca's tragedies were characterized by ghosts demanding revenge, themes of blood guilt, the rhetoric of intense moral deliberation, and a pervasive sense of doom. Shakespeare adopts all of these. The Ghost of King Hamlet is directly Senecan: a spirit from the underworld commanding the living to exact vengeance. The Ghost declares, "), ...cite("Revenge his foul and most unnatural murder", "Act I", "Scene 5", "line 25"), t(", framing the entire dramatic action in Senecan terms of blood obligation. Unlike the Greek model, which tends toward fate and divine machinery, Senecan drama is driven by personal passion, family honour, and the spectacle of moral corruption — all present in "), italic("Hamlet.")]),
  para([t("The Senecan element is also visible in the play's preoccupation with interior psychological states. Seneca's characters frequently deliver long, anguished speeches of self-analysis — a device Shakespeare elevates into the soliloquy. Claudius's agonised prayer — "), ...cite("O, my offence is rank, it smells to heaven; / It hath the primal eldest curse upon't, / A brother's murder", "Act III", "Scene 3", "lines 36–38"), t(" — is quintessentially Senecan in its tortured self-revelation. Furthermore, the play's bloodbath finale, where Gertrude, Laertes, Claudius, and Hamlet all die within moments of each other, mirrors the spectacular carnage of Senecan drama, where justice is achieved only through universal destruction.")]),

  heading2("English (Elizabethan) Tragic Elements"),
  para([t("Shakespeare transforms his classical inheritance by embedding "), italic("Hamlet"), t(" within a distinctly English political and moral framework. The Elizabethan tragic tradition was shaped by the "), italic("de casibus"), t(" formula — the fall of great men from Fortune's wheel — and by the morality play tradition that demanded clear ethical lessons. Shakespeare complicates this by refusing easy moral resolution. Hamlet is simultaneously avenger and murderer, philosopher and coward, lover and misogynist. The play engages with specifically English concerns: the legitimacy of monarchy, the corruption of the court, the Protestant theological anxieties about purgatory and the afterlife.")]),
  para([t("The Ghost's ambiguous status — is it a genuine spirit or a devil in disguise? — reflects the post-Reformation English anxiety about Catholic doctrines of purgatory. Hamlet acknowledges this uncertainty: "), ...cite("The spirit that I have seen / May be the devil", "Act II", "Scene 2", "lines 598–599"), t(". This theological hesitation is peculiarly English and Protestant, introducing a layer of epistemological doubt absent from Greek or Roman tragedy. The political dimension — Denmark as a mirror for Elizabethan England — adds a further English specificity, as the rottenness at court reflects anxieties about succession, counsel, and the body politic that were deeply topical in Shakespeare's England.")]),
];

// ============================================================
// THEME 2: Hamlet as Ideal Tragic Hero
// ============================================================
const theme2 = () => [
  pageBreak(),
  heading1("Theme 2: Hamlet — The Tragic Sense (As an Ideal Tragic Hero)"),
  thesis("Hamlet embodies the ideal tragic hero not by virtue of conventional heroic action but through his extraordinary capacity for moral consciousness, which simultaneously elevates him above ordinary humanity and renders him incapable of the decisive violence that conventional revenge demands."),

  heading2("Nobility and the Heroic Endowment"),
  para([t("Aristotle's requirement that the tragic hero be a man of noble birth and character is amply satisfied by Hamlet. Ophelia's celebrated eulogy captures what Hamlet was before his father's murder disrupted his world: "), ...cite("The courtier's, soldier's, scholar's, eye, tongue, sword, / The expectancy and rose of the fair state, / The glass of fashion and the mould of form, / The observed of all observers", "Act III", "Scene 1", "lines 154–157"), t(". This description locates Hamlet at the intersection of all the values Elizabethan culture prized: courtly grace, martial capability, and intellectual distinction. His education at Wittenberg aligns him with the new humanist learning of the Renaissance, while his sensitivity to music, theatre, and philosophy marks him as a man of genuine refinement. Shakespeare takes pains to establish this before showing his disintegration, making the gap between what Hamlet was and what the revenge plot demands of him the source of the play's tragic pathos.")]),
  para([t("The concept of "), italic("hamartia"), t(" — the tragic flaw — has been debated extensively in relation to Hamlet. Unlike Macbeth's ambition or Othello's jealousy, Hamlet's flaw is not a moral vice but an intellectual excess: his tendency to think too precisely on an event. He himself diagnoses this: "), ...cite("Thus conscience does make cowards of us all, / And thus the native hue of resolution / Is sicklied o'er with the pale cast of thought", "Act III", "Scene 1", "lines 82–84"), t(". The word 'conscience' here carries its Renaissance double meaning — both moral scruple and reflective consciousness. It is Hamlet's capacity for both that constitutes his flaw: he cannot act without fully understanding the moral implications of his action, and in the moral complexity of the revenge situation, full understanding is impossible.")]),

  heading2("The Tragic Fall and Peripeteia"),
  para([t("The ideal tragic hero must undergo a reversal of fortune — "), italic("peripeteia"), t(" — that is both surprising and inevitable. Hamlet's downfall unfolds with inexorable logic despite his superior intelligence. Every action he takes to advance his revenge deepens his entrapment. His decision to stage 'The Mousetrap' to confirm Claudius's guilt is brilliant strategy, but it alerts Claudius to Hamlet's knowledge and sets the counter-plot in motion. His failure to kill Claudius in prayer "), ...cite("Now might I do it pat, now he is praying; / And now I'll do't. And so he goes to heaven; / And so am I revenged", "Act III", "Scene 3", "lines 73–75"), t(" — though rationalised by theological reasoning — represents the tragic missed opportunity that seals his fate. The accidental killing of Polonius, executed in impulsive haste rather than deliberate revenge, is the fulcrum of the play's peripeteia: it transforms Hamlet from avenger to murderer and precipitates the catastrophe.")]),
  para([t("The final duel scene achieves Aristotelian "), italic("catharsis"), t(" with remarkable theatrical power. Hamlet fulfils his revenge, kills Claudius, but at the cost of his own life. His last words carry the dignity and philosophical resignation of the ideal tragic hero: "), ...cite("The rest is silence", "Act V", "Scene 2", "line 363"), t(". Horatio's tribute reinforces this: "), ...cite("Now cracks a noble heart. Good night, sweet prince, / And flights of angels sing thee to thy rest", "Act V", "Scene 2", "lines 370–371"), t(". This elegy, one of the most beautiful in English literature, confirms Hamlet's status as a genuinely noble figure whose destruction the audience mourns — the essence of tragic catharsis. The arrival of Fortinbras restores political order, but the loss of Hamlet represents something irreplaceable: a kind of moral and intellectual greatness that cannot be reproduced.")]),

  heading2("Psychological Complexity as Tragic Depth"),
  para([t("What distinguishes Hamlet as an ideal tragic hero for modern audiences is precisely what puzzled earlier critics: his psychological complexity. Unlike the more straightforward tragic heroes of Greek drama, Hamlet possesses an interior life of almost inexhaustible depth, revealed through seven soliloquies that take the audience inside his consciousness with unprecedented intimacy. The 'To be or not to be' soliloquy "), ...cite("To be, or not to be, that is the question: / Whether 'tis nobler in the mind to suffer / The slings and arrows of outrageous fortune, / Or to take arms against a sea of troubles", "Act III", "Scene 1", "lines 55–58"), t(" is not merely a personal meditation on suicide but a philosophical examination of the human condition. By universalising Hamlet's dilemma through the pronoun 'we', Shakespeare makes his hero's tragedy representative of all humanity's confrontation with suffering and mortality.")]),
];

// ============================================================
// THEME 3: Madness and Procrastination
// ============================================================
const theme3 = () => [
  pageBreak(),
  heading1("Theme 3: The Theme of Madness and Procrastination in Hamlet"),
  thesis("In Hamlet, madness functions as both a strategic mask and a psychological reality, while procrastination emerges not as weakness but as the inevitable consequence of a consciousness too morally refined to execute violence without absolute certainty — together, these twin themes illuminate Shakespeare's central argument that thought and action are fundamentally incompatible in a corrupt world."),

  heading2("Feigned Madness as Strategic Tool"),
  para([t("Hamlet's decision to adopt an 'antic disposition' is announced immediately after his encounter with the Ghost: "), ...cite("How strange or odd soe'er I bear myself, / As I perchance hereafter shall think meet / To put an antic disposition on", "Act I", "Scene 5", "lines 170–172"), t(". This feigned madness serves multiple strategic functions. It provides cover for his investigation of Claudius's guilt, as a mad prince is politically non-threatening. It allows him to speak truths that would be dangerous from a sane man — his cutting remarks to Polonius, his brutal honesty to Ophelia, his barbed commentary at the play-within-a-play are all sheltered under the guise of madness. As his course materials note, Hamlet's language under the 'antic disposition' is 'ambiguous and philosophical', allowing him to communicate double meanings that a sane man could not safely utter at court.")]),
  para([t("However, the boundary between feigned and real madness becomes increasingly unstable as the play progresses. Hamlet's behaviour becomes genuinely erratic — his savage rejection of Ophelia, "), ...cite("Get thee to a nunnery, why wouldst thou be a breeder of sinners?", "Act III", "Scene 1", "lines 121–122"), t(", his impulsive killing of Polonius, his ghoulish humour with the gravediggers — all suggest a man whose psychological equilibrium has been genuinely disturbed. The court is uncertain, and so is the audience. This ambiguity is deliberately maintained by Shakespeare as a reflection of the play's broader epistemological theme: the impossibility of distinguishing appearance from reality in Elsinore.")]),

  heading2("Real Madness: Ophelia's Tragedy"),
  para([t("Shakespeare places Ophelia's genuine madness in deliberate contrast to Hamlet's feigned disorder. While Hamlet's 'madness' is purposeful and controlled, Ophelia's collapse in Act IV is real, unmediated, and heartbreaking. Having lost her father to Hamlet's violence and been rejected by the man she loved, she enters singing fragmentary songs and distributing flowers with a symbolic precision that her sane mind could not have articulated. Her distribution of rosemary for remembrance, pansies for thoughts, rue for sorrow, and fennel for flattery constitutes a devastatingly accurate moral commentary on the court: "), ...cite("There's rosemary, that's for remembrance; pray you, love, remember", "Act IV", "Scene 5", "line 174"), t(". Ophelia's madness reveals what Hamlet's feigned madness conceals: the true emotional cost of Elsinore's corruption.")]),
  para([t("The contrast between the two forms of madness serves Shakespeare's thematic purpose precisely. Hamlet uses madness as a tool of intelligence; Ophelia is destroyed by the actions of those who wielded power over her. His madness has method, as Polonius grudgingly acknowledges: "), ...cite("Though this be madness, yet there is method in't", "Act II", "Scene 2", "lines 205–206"), t(". Hers has none — it is pure emotional devastation. Together, the two madnesses form a comprehensive portrait of how an unjust political order distorts and destroys human consciousness.")]),

  heading2("Procrastination: Cause and Consequence"),
  para([t("Hamlet's delay in avenging his father has generated more critical commentary than almost any other feature of the play. The delay is not simple cowardice — Hamlet demonstrates physical courage on the sea voyage and in the duel — but rather the consequence of his extraordinary moral and intellectual constitution. He cannot act without certainty of guilt (hence 'The Mousetrap'), without theological clarity about the consequences (hence the prayer scene hesitation), and without resolving his philosophical ambivalence about whether human life — even a murderer's life — should be taken. His seventh soliloquy "), ...cite("How all occasions do inform against me, / And spur my dull revenge!", "Act IV", "Scene 4", "lines 32–33"), t(" acknowledges the delay but does not overcome it.")]),
  para([t("The contrast with Fortinbras illuminates the nature of Hamlet's procrastination most sharply. Fortinbras leads an army to fight and die for a worthless patch of ground — pure action without reflection. Hamlet admires and is shamed by this, yet simultaneously recognises its absurdity: "), ...cite("Rightly to be great / Is not to stir without great argument, / But greatly to find quarrel in a straw / When honour's at the stake", "Act IV", "Scene 4", "lines 53–56"), t(". The bitter irony is that Hamlet has the greatest possible argument for action — a murdered father, a usurped throne, a corrupted kingdom — yet cannot act, while Fortinbras acts for nothing. This juxtaposition suggests that procrastination is not merely a personal failing but a structural consequence of the tragic hero's consciousness: to think deeply is to be paralysed by complexity.")]),
];

// ============================================================
// THEME 4: Soliloquies
// ============================================================
const theme4 = () => [
  pageBreak(),
  heading1("Theme 4: The Use of Soliloquies in Hamlet"),
  thesis("Shakespeare's seven soliloquies in Hamlet constitute the most sustained experiment in dramatic interiority in the Western theatrical tradition, transforming the revenge play's conventional focus on external action into a profound psychological drama in which consciousness itself becomes both the subject and the site of tragedy."),

  heading2("The Function and Form of the Soliloquy"),
  para([t("The soliloquy as a dramatic device predates Shakespeare, but no playwright before him exploited it with such psychological intensity and thematic purposefulness. In "), italic("Hamlet"), t(", the soliloquies serve as windows into a consciousness that the action of the play cannot fully reveal. They establish dramatic irony — the audience knows truths that other characters do not — and they create an intimacy between Hamlet and the spectator that is unique in Renaissance drama. Shakespeare's language in the soliloquies is characteristically dense with metaphor, syntactic complexity, and tonal variation, reflecting the gyrations of a powerful intellect under emotional stress.")]),
  para([t("The first soliloquy, "), ...cite("O, that this too, too solid flesh would melt, / Thaw, and resolve itself into a dew!", "Act I", "Scene 2", "lines 129–130"), t(", establishes the template for all that follow. Hamlet wishes for dissolution — not death exactly (suicide is theologically forbidden) but escape from the intolerable burden of consciousness in a corrupt world. The imagery of an 'unweeded garden / That grows to seed' "), ...cite("Tis an unweeded garden / That grows to seed; things rank and gross in nature / Possess it merely", "Act I", "Scene 2", "lines 135–137"), t(" introduces the play's pervasive imagery of disease and natural corruption, and the soliloquy's sudden shift from cosmic despair to bitterly specific grievance against Gertrude demonstrates the rapid modulations of feeling that Shakespeare will sustain throughout.")]),

  heading2("The Central Soliloquies and Their Philosophical Stakes"),
  para([t("The third soliloquy, delivered after the Players' arrival, represents a crucial turning point in Hamlet's self-understanding. Watching the Player weep for Hecuba — a fictional character — Hamlet is shamed by his own inaction: "), ...cite("What's Hecuba to him, or he to Hecuba, / That he should weep for her? What would he do / Had he the motive and the cue for passion / That I have?", "Act II", "Scene 2", "lines 559–562"), t(". This comparison of theatrical emotion to genuine moral obligation is philosophically rich: the actor who can simulate passion for fiction exposes Hamlet's inability to translate genuine passion into action. The soliloquy ends with the plan for 'The Mousetrap' — a rare moment of decisive strategy — but even this action is mediated through theatre, through representation rather than direct confrontation.")]),
  para([t("The fourth soliloquy — 'To be or not to be' — is the pinnacle of Shakespeare's achievement in the form. It is unique among the seven in that Hamlet speaks not personally but universally, using 'we' and 'us' to expand his individual dilemma into a meditation on the human condition. The question is not simply whether to commit suicide but whether to resist the suffering of existence or to end it, and what prevents this ending: "), ...cite("The undiscovered country, from whose bourn / No traveller returns, puzzles the will, / And makes us rather bear those ills we have / Than fly to others that we know not of", "Act III", "Scene 1", "lines 78–81"), t(". The 'undiscovered country' of death — a magnificent phrase — represents not just the afterlife but all the unknowns that paralyse human action. This generalisation transforms Hamlet's particular situation into a universal existential statement.")]),

  heading2("Soliloquies as Dramatic Architecture"),
  para([t("The seven soliloquies are not merely character revelations but structural pillars of the play's dramatic architecture. They mark the stages of Hamlet's psychological journey: from initial depression (Soliloquy 1) through shock and vows of revenge (Soliloquy 2), self-criticism and strategic planning (Soliloquy 3), universal philosophical meditation (Soliloquy 4), violent resolution restrained by moral scruple (Soliloquy 5), theological justification for further delay (Soliloquy 6), to the shame-driven 'bloody thoughts' resolution (Soliloquy 7). This arc — depression to philosophical acceptance — constitutes the inner action of the play, the story of a consciousness rather than merely a plot.")]),
  para([t("The final soliloquy, delivered after observing Fortinbras's army march to war, reaches its climax with a declaration that suggests resolution but maintains ambiguity: "), ...cite("O, from this time forth, / My thoughts be bloody, or be nothing worth!", "Act IV", "Scene 4", "lines 65–66"), t(". The resolution is expressed in terms of 'thoughts' rather than actions — a telling distinction. Even here, at the moment of apparent commitment, Hamlet's characteristic habit of converting action into thought reasserts itself. The soliloquy form, which by its nature is thought rather than action, perfectly embodies this irresolvable tension. It is ultimately the soliloquies themselves — thought given form — that Shakespeare presents as Hamlet's most authentic mode of being.")]),
];

// ============================================================
// THEME 5: Death
// ============================================================
const theme5 = () => [
  pageBreak(),
  heading1("Theme 5: The Theme of Death (The Physical and the Spiritual)"),
  thesis("In Hamlet, death operates simultaneously on physical, philosophical, and theological registers, functioning not merely as dramatic outcome but as the play's central epistemological problem — the unknowable 'undiscovered country' whose uncertainty generates the entire tragic action."),

  heading2("Physical Death and Its Spectacle"),
  para([t("Shakespeare presents death with unusual physical directness in "), italic("Hamlet"), t(". The graveyard scene in Act V is the most sustained meditation on physical mortality in all of Shakespeare. The gravedigger's casual professionalism — digging graves and tossing up skulls with cheerful indifference — creates a powerful memento mori. When Hamlet takes up Yorick's skull, the physical reality of death — the reduction of a beloved person to dry bone — is made viscerally present: "), ...cite("Alas, poor Yorick! I knew him, Horatio — a fellow of infinite jest, of most excellent fancy... Where be your gibes now? Your gambols? Your songs?", "Act V", "Scene 1", "lines 184–188"), t(". This speech moves from affectionate personal memory to sardonic contemplation of universal decay, as Hamlet imagines even Alexander the Great's remains stopping a bunghole. Death is the great leveller, reducing noble and base alike to dust.")]),
  para([t("The play's pattern of deaths is also physically significant. Each death in "), italic("Hamlet"), t(" is in some sense wrongful or accidental: King Hamlet is murdered; Polonius is killed by mistake; Ophelia's death is ruled accidental by a sympathetic coroner, though it appears suicidal; Rosencrantz and Guildenstern are executed by Hamlet's manipulation of their commission; Gertrude drinks poison meant for Hamlet; Laertes is killed by his own poisoned sword; Claudius dies by both sword and poison; Hamlet himself dies from a poisoned wound. This accumulation of deaths — all interconnected through Claudius's original murder — suggests that crime against nature generates an expanding circle of destruction that cannot be contained.")]),

  heading2("The Spiritual Dimension of Death"),
  para([t("The theological dimensions of death in "), italic("Hamlet"), t(" are among the play's most carefully developed concerns. The Ghost's account of his own death introduces immediately the question of what lies beyond physical dissolution. He is condemned to walk the night and suffer purgatorial purification during the day: "), ...cite("I am thy father's spirit, / Doom'd for a certain term to walk the night, / And for the day confin'd to fast in fires, / Till the foul crimes done in my days of nature / Are burnt and purg'd away", "Act I", "Scene 5", "lines 9–13"), t(". This explicitly Catholic doctrine of purgatory would have been theologically charged for Shakespeare's Protestant audience, for whom purgatory had been officially abolished. The Ghost's existence is therefore not only literally uncanny but theologically ambiguous.")]),
  para([t("Hamlet's own famous meditation on death in the 'To be or not to be' soliloquy centres on the spiritual uncertainty: what happens after death? The fear of 'the undiscovered country, from whose bourn / No traveller returns' "), ...cite("But that the dread of something after death, / The undiscovered country, from whose bourn / No traveller returns, puzzles the will", "Act III", "Scene 1", "lines 77–79"), t(" is precisely this: we cannot know. This epistemological void — not pain but uncertainty — is what makes humans endure life's sufferings rather than seek release. Hamlet's delay in the prayer scene is also spiritually motivated: he refuses to kill Claudius while he prays because that might send him to heaven, and "), ...cite("that would be scann'd: / A villain kills my father; and for that, / I, his sole son, do this same villain send / To heaven", "Act III", "Scene 3", "lines 75–78"), t(". The irony — which Shakespeare certainly intends — is that Claudius's prayer is ineffective, so Hamlet's theological scruple prevents what would actually have been perfectly timed revenge.")]),

  heading2("Death as Resolution and Acceptance"),
  para([t("Hamlet's final attitude toward death represents the play's most significant spiritual development. Having returned from England with a new equanimity — "), ...cite("There is a special providence in the fall of a sparrow. If it be now, 'tis not to come; if it be not to come, it will be now; if it be not now, yet it will come — the readiness is all", "Act V", "Scene 2", "lines 219–222"), t(" — Hamlet has moved from fear of death to philosophical acceptance. The phrase 'the readiness is all' echoes King Lear's 'Ripeness is all' and represents a distinctly Shakespearean tragic wisdom: the understanding that human beings cannot control the timing or manner of their death, but can control their readiness to meet it. This acceptance is not resignation but a kind of spiritual maturity achieved through suffering.")]),
];

// ============================================================
// THEME 6: Role of Women
// ============================================================
const theme6 = () => [
  pageBreak(),
  heading1("Theme 6: The Role of Women in Hamlet"),
  thesis("Shakespeare uses Gertrude and Ophelia not as mere supporting characters but as structural mirrors that reflect the patriarchal violence at the heart of Elsinore, their constrained agency and tragic fates constituting a devastating implicit critique of a social order that destroys what it claims to protect."),

  heading2("Gertrude: Agency, Ambiguity, and Misrepresentation"),
  para([t("Gertrude is the most contested character in "), italic("Hamlet"), t(". The text provides genuinely ambiguous evidence about her moral culpability: did she know of Claudius's murder? Did she love him or marry him out of political necessity? Shakespeare refuses to resolve these questions, and this ambiguity is itself significant. Hamlet's furious indictment — "), ...cite("O, most wicked speed, to post / With such dexterity to incestuous sheets!", "Act I", "Scene 2", "lines 156–157"), t(" — tells us more about Hamlet's misogynistic disgust than about Gertrude's actual motivations. His generalisation from his mother's conduct to all women — 'Frailty, thy name is woman!' "), ...cite("Frailty, thy name is woman!", "Act I", "Scene 2", "line 146"), t(" — is a classic instance of patriarchal reasoning: the behaviour of one woman becomes an indictment of all women.")]),
  para([t("In the closet scene, Gertrude's response to Hamlet's accusations suggests genuine remorse: "), ...cite("O Hamlet, thou hast cleft my heart in twain", "Act III", "Scene 4", "line 157"), t(". Whether this represents genuine moral awakening or simply capitulation to her son's emotional pressure is deliberately unclear. What is clear is that Gertrude's position throughout the play is one of extreme constraint. She is queen, but her power derives entirely from her relationship to men — first her husband, then his brother, and in a different sense her son. She has no autonomous political identity. Her final act — drinking from the poisoned cup despite Claudius's warning — may be read as suicide, maternal sacrifice, or tragic ignorance. This interpretive openness is itself a comment on the difficulty of understanding women's inner lives through a narrative entirely controlled by male perspectives.")]),

  heading2("Ophelia: Innocence, Constraint, and Destruction"),
  para([t("Ophelia's tragedy is, in some ways, more complete than Hamlet's because she lacks even his limited capacity for resistance. She is subject to the commands of her father Polonius, the manipulation of King Claudius, and the emotional violence of Hamlet. Polonius's instruction to reject Hamlet's love letters and deny him access is presented as paternal protection, but it functions as political instrumentalisation: "), ...cite("You do not understand yourself so clearly / As it behoves my daughter and your honour", "Act I", "Scene 3", "lines 96–97"), t(". Ophelia obeys, and this obedience — which she has been trained to regard as virtue — contributes directly to Hamlet's destabilisation and her own destruction.")]),
  para([t("Hamlet's treatment of Ophelia in the 'nunnery scene' represents a concentrated instance of the play's patriarchal violence. His instruction "), ...cite("Get thee to a nunnery, go. Farewell. Or, if thou wilt needs marry, marry a fool, for wise men know well enough what monsters you make of them", "Act III", "Scene 1", "lines 137–140"), t(" is ostensibly a warning but functions as a brutal rejection that destroys her sense of self. His subsequent staging of cruelty toward her during 'The Mousetrap' — making sexual jests in public — completes the humiliation. Ophelia's madness in Act IV is the direct consequence of being used, instrumentalised, and then discarded by every man in her life. Her flower distribution, her fragmentary songs, her drowning — all constitute a form of protest that her sane, controlled self could never have expressed. In death and madness, Ophelia paradoxically achieves her most authentic self-expression.")]),
];

// ============================================================
// THEME 7: Symbols and Style of Shakespeare
// ============================================================
const theme7 = () => [
  pageBreak(),
  heading1("Theme 7: The Use of Symbols and the Style of Shakespeare"),
  thesis("Shakespeare's symbolic vocabulary in Hamlet — centred on images of disease, decay, performance, and corruption — works in seamless conjunction with his distinctive stylistic resources of blank verse, soliloquy, and dramatic irony to produce a work in which form and meaning are inseparable, where the style is the statement."),

  heading2("Central Symbols: Disease, Decay, and Corruption"),
  para([t("The dominant symbolic pattern of "), italic("Hamlet"), t(" is the imagery of disease, infection, and physical decay. The play opens with a rottenness in the state of Denmark that is more than metaphorical — it reflects a genuine moral and political corruption emanating from Claudius's fratricide and usurpation. Marcellus's famous observation "), ...cite("Something is rotten in the state of Denmark", "Act I", "Scene 4", "line 90"), t(" establishes the controlling metaphor immediately. Claudius's crime against his brother — pouring poison into his ear — is literally a corruption of the body, and this physical detail radiates outward into the play's pervasive imagery of festering, ulcers, and infection. Hamlet tells Gertrude: "), ...cite("It will but skin and film the ulcerous place, / Whilst rank corruption, mining all within, / Infects unseen", "Act III", "Scene 4", "lines 147–149"), t(". The image perfectly captures the play's political argument: evil that is concealed rather than confronted continues to spread invisibly.")]),
  para([t("The Ghost is itself a powerful symbol — of violated natural order, of the past's claim on the present, of the unresolved crimes that haunt political communities. The use of the theatre-within-the-theatre — 'The Mousetrap' — is both a plot device and a complex symbol of art's capacity to reveal truth that ordinary reality conceals. Hamlet's instruction to the Players "), ...cite("the purpose of playing... was and is, to hold, as 'twere, the mirror up to nature", "Act III", "Scene 2", "lines 20–22"), t(" articulates Shakespeare's own theory of dramatic art: theatre as truth-telling, as reality-testing.")]),

  heading2("Style: Blank Verse, Prose, and the Dramatic Register"),
  para([t("Shakespeare's stylistic virtuosity in "), italic("Hamlet"), t(" is manifested in his flexible deployment of blank verse and prose to reflect character and psychological state. The play's principal characters speak predominantly in blank verse — the iambic pentameter that Shakespeare perfected as the vehicle for elevated dramatic speech. But Hamlet's shifts into prose during his feigned madness, his conversations with the Players, and his exchanges with the gravediggers are deliberately stylistically marked: prose signals either social intimacy, comic register, or the collapse of conventional order.")]),
  para([t("Within the verse itself, Shakespeare achieves extraordinary effects through metrical variation. The famous opening of the 'To be or not to be' soliloquy violates regular iambic rhythm to enact Hamlet's psychological stress, while Claudius's smooth, rhetorically polished addresses to the court reveal his political calculation through their very formal eloquence. The contrast between Hamlet's fractured, hyperactive imagery and Claudius's controlled periodic sentences is a constant stylistic commentary on the difference between authentic feeling and political performance. Even Polonius's famous apothegm "), ...cite("This above all: to thine own self be true, / And it must follow, as the night the day, / Thou canst not then be false to any man", "Act I", "Scene 3", "lines 78–80"), t(" is linguistically ironic: these wise words come from a man who is himself fundamentally false, and the formal sententious style enacts this hypocrisy.")]),

  heading2("Dramatic Irony and the Theatrical Self-Consciousness"),
  para([t("Shakespeare's style in "), italic("Hamlet"), t(" is also characterised by a theatrical self-consciousness unusual even for Renaissance drama. The play-within-the-play is the most obvious instance, but there are others: Hamlet's instructions to the Players about naturalistic acting, his numerous theatrical metaphors ('all the world's a stage' tradition), and his role as both director and performer within the revenge drama he is simultaneously experiencing. The audience watching "), italic("Hamlet"), t(" watch Hamlet watching the play that mirrors reality — a hall of mirrors of theatrical consciousness. This reflects what Mufeed Al Abdullah and Susane Ramadan have identified as Shakespeare's 'language strategies': his use of multilayered discourse that operates on several registers simultaneously, so that a single speech communicates different things to different hearers within the play and to the audience in the theatre.")]),
];

// ============================================================
// THEME 8: Tripartite World View
// ============================================================
const theme8 = () => [
  pageBreak(),
  heading1("Theme 8: The Tripartite World View of the Elizabethan Age as Expressed in Hamlet's Speeches"),
  thesis("Hamlet's speeches embody the Elizabethan tripartite world view — the Chain of Being, the macrocosm-microcosm correspondence, and the body politic — not as static intellectual inheritance but as a living tension between the inherited medieval cosmology and the emerging humanist scepticism of the Renaissance, a tension that generates the play's deepest philosophical conflicts."),

  heading2("The Great Chain of Being and Its Disruption"),
  para([t("The Elizabethan world view, as articulated by scholars including Mufeed Al Abdullah and Susane Ramadan in "), italic("Shakespeare's Language Strategies"), t(", was organised around a hierarchical conception of the universe known as the Great Chain of Being. This cosmological framework arranged all existence in a divinely ordered hierarchy from God through the angels, the celestial spheres, kings, nobility, common people, animals, and down to inanimate matter. Crucially, disruption at any level of this chain was believed to produce sympathetic disorder at all other levels — political murder would be reflected in cosmic disturbance, natural monstrosity, and social chaos. "), italic("Hamlet"), t(" explicitly invokes this framework. The Ghost tells Hamlet that Claudius has committed a crime that has disrupted nature itself: "), ...cite("Murder most foul, as in the best it is; / But this most foul, strange, and unnatural", "Act I", "Scene 5", "lines 27–28"), t(". The word 'unnatural' carries its full Elizabethan weight: Claudius's fratricide and usurpation violate not merely human law but the divinely ordained natural order.")]),
  para([t("Hamlet's first soliloquy registers this cosmic disorder through the image of the world as an 'unweeded garden': "), ...cite("Tis an unweeded garden / That grows to seed; things rank and gross in nature / Possess it merely", "Act I", "Scene 2", "lines 135–137"), t(". In Elizabethan symbolic thinking, the ordered garden was a principal emblem of the properly maintained Chain of Being; its corruption into rank weeds signals precisely the disorder that Claudius's crime has unleashed. Hamlet's celebrated address to the heavens further invokes this tripartite world view: "), ...cite("What a piece of work is a man! How noble in reason, how infinite in faculty! In form and moving how express and admirable! In action how like an angel! In apprehension how like a god!", "Act II", "Scene 2", "lines 303–306"), t(". This speech positions humanity within the Elizabethan Chain — above the beasts in reason, below the angels in spiritual purity — while its bitter coda ('And yet, to me, what is this quintessence of dust?') signals Hamlet's crisis of faith in this inherited framework.")]),

  heading2("The Macrocosm-Microcosm Correspondence"),
  para([t("A second pillar of the Elizabethan tripartite world view was the doctrine of correspondence between the macrocosm (the universe) and the microcosm (the human body and the individual state). As Al Abdullah and Ramadan have shown in their analysis of Shakespeare's language strategies, Shakespeare consistently exploits these correspondences to generate political meaning. In "), italic("Hamlet"), t(", the corruption of the state of Denmark corresponds precisely to the corruption of the human body. Claudius's original crime — pouring poison into King Hamlet's ear — operates simultaneously on the physical, political, and cosmic registers. The Ghost describes the physical effects in gruesome detail: "), ...cite("And in the porches of my ears did pour / The leperous distilment... and a most instant tetter bark'd about, / Most lazar-like, with vile and loathsome crust, / All my smooth body", "Act I", "Scene 5", "lines 63–73"), t(". This physical corruption mirrors the political corruption Claudius introduces — the body politic infected by the same poison as the body natural.")]),
  para([t("Hamlet's celebrated address to Denmark as a 'prison' — "), ...cite("Denmark's a prison", "Act II", "Scene 2", "line 243"), t(" — operates within this framework of correspondence. A nation ruled by a fratricide becomes literally a place of confinement for all its inhabitants, because the corruption at the head infects the whole body politic. The macrocosmic disorder (fratricidal usurpation) is reflected in microcosmic disorders: the instability of individual identities (who can be trusted?), the breakdown of family bonds (Hamlet's relationship with Gertrude, Ophelia's rejection by Hamlet), and the collapse of social roles (a king who is a murderer, a father who spies on his daughter).")]),

  heading2("The Renaissance Challenge to the Tripartite Order"),
  para([t("What makes "), italic("Hamlet"), t(" philosophically extraordinary is that it does not simply dramatise the Elizabethan tripartite world view — it simultaneously questions and undermines it. Hamlet's education at Wittenberg aligns him with the new Renaissance humanism that was challenging the old cosmological certainties. His scepticism about the Ghost — "), ...cite("The spirit that I have seen / May be the devil, and the devil hath power / To assume a pleasing shape", "Act II", "Scene 2", "lines 598–600"), t(" — represents a genuinely modern epistemological caution: the traditional world view says ghosts from purgatory can demand revenge; Hamlet's humanist training requires empirical verification before he trusts any such supernatural claim. This tension between inherited cosmology and emerging empiricism is precisely what Al Abdullah and Ramadan identify as central to Shakespeare's language strategies: he uses the tripartite framework not as stable ideology but as a site of contestation, where old certainties are perpetually interrogated by new forms of knowledge.")]),
  para([t("Hamlet's final acceptance — 'the readiness is all' "), ...cite("If it be now, 'tis not to come; if it be not to come, it will be now; if it be not now, yet it will come. The readiness is all", "Act V", "Scene 2", "lines 220–222"), t(" — represents a resolution to this tension that is neither the old providential world view nor pure humanist scepticism, but something new: a tragic wisdom that accepts the limits of human understanding while refusing to abandon moral responsibility. Shakespeare, through Hamlet's speeches, enacts the intellectual transition of the Elizabethan age itself — from the ordered certainty of the medieval Chain of Being to the anxious, questioning modernity that would eventually produce the scientific revolution and the Enlightenment.")]),
];

// Build document
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Times New Roman", color: "1F3864" },
        paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Times New Roman", color: "2E5B9A" },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Hamlet Thematic Analysis  |  Page ", size: 20, font: "Times New Roman", color: "666666" }),
            new PageNumber({ size: 20, font: "Times New Roman", color: "666666" })
          ]
        })]
      })
    },
    children: [
      ...titlePage(),
      ...theme1(),
      ...theme2(),
      ...theme3(),
      ...theme4(),
      ...theme5(),
      ...theme6(),
      ...theme7(),
      ...theme8(),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/home/claude/Hamlet_Thematic_Analysis.docx', buffer);
  console.log('Done! File created.');
}).catch(err => {
  console.error('Error:', err);
  process.exit(1);
});
        
