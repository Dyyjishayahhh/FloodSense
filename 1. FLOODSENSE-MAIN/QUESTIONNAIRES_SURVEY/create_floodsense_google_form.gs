/**
 * Creates the complete FloodSense Resident Survey as a Google Form.
 *
 * How to use:
 * 1. Open https://script.google.com and create a new project.
 * 2. Paste this entire file into Code.gs and save it.
 * 3. Select createFloodSenseForm from the function list, then click Run.
 * 4. Approve the requested Google Forms permission.
 * 5. Open the execution log to get the editor and respondent URLs.
 *
 * Run the function only once unless you intentionally want another copy.
 */
function createFloodSenseForm() {
  const form = FormApp.create('FloodSense Resident Survey Questionnaire', true);

  form
    .setDescription(
      'We are studying how Bacoor City residents receive and use flood-related information. ' +
        'We want to understand what currently works, what problems residents experience, ' +
        'and whether a proposed FloodSense application could be useful.\n\n' +
        'Pinag-aaralan namin kung paano natatanggap at ginagamit ng mga residente ng Bacoor City ' +
        'ang impormasyon tungkol sa baha. Nais naming malaman kung ano ang kasalukuyang gumagana, ' +
        'anong mga problema ang nararanasan, at kung maaaring makatulong ang iminungkahing ' +
        'FloodSense application.\n\n' +
        'Participation is voluntary. Please review the privacy notice and indicate your consent ' +
        'before continuing. Select only one answer unless the question says otherwise.\n\n' +
        'Kusang-loob ang pagsali. Basahin ang pahayag sa pagkapribado at ipahayag ang inyong ' +
        'pahintulot bago magpatuloy. Pumili lamang ng isang sagot maliban kung may ibang panuto.'
    )
    .setConfirmationMessage(
      'Thank you. Your response has been recorded. / Salamat. Naitala na ang inyong sagot.'
    )
    .setProgressBar(true)
    .setShuffleQuestions(false)
    .setShowLinkToRespondAgain(false)
    .setIsQuiz(false);

  form
    .addSectionHeaderItem()
    .setTitle('Data Privacy Notice and Consent / Pahayag sa Pagkapribado at Pahintulot')
    .setHelpText(
      'In accordance with Republic Act No. 10173, or the Data Privacy Act of 2012, the ' +
        'FloodSense research team will collect and process the personal information you provide, ' +
        'including your name, age, sex or gender, and general place of residence, only for this ' +
        'academic study on flood-information needs in Bacoor City. Participation is voluntary. ' +
        'You may skip a demographic question or stop answering at any time without penalty. ' +
        'Your responses will be kept confidential, accessed only by the research team and its ' +
        'academic advisers, reported in summary form whenever possible, retained only as long as ' +
        'necessary for the study and required academic recordkeeping, and then securely disposed ' +
        'of. Your information will not be sold or used for marketing. Subject to applicable law, ' +
        'you may request access to, correction of, or deletion of your information, or withdraw ' +
        'your consent, by contacting the FloodSense research team through the survey administrator.\n\n' +
        'Alinsunod sa Republic Act No. 10173, o Data Privacy Act of 2012, kokolektahin at ' +
        'gagamitin ng FloodSense research team ang personal na impormasyong ibibigay ninyo, ' +
        'kabilang ang inyong pangalan, edad, kasarian, at pangkalahatang lugar ng tirahan, para ' +
        'lamang sa akademikong pag-aaral na ito tungkol sa pangangailangan sa impormasyong may ' +
        'kaugnayan sa baha sa Bacoor City. Kusang-loob ang pagsali. Maaari ninyong laktawan ang ' +
        'isang tanong sa demograpiko o huminto anumang oras nang walang kaparusahan. Pananatiling ' +
        'kumpidensiyal ang inyong mga sagot, at tanging ang research team at mga academic adviser ' +
        'nito ang magkakaroon ng access. Ilalahad ang mga resulta sa pinagsama-samang paraan ' +
        "hangga't maaari, itatago lamang ang datos habang kailangan para sa pag-aaral at sa " +
        'kinakailangang academic recordkeeping, at pagkatapos ay ligtas na buburahin o sisirain. ' +
        'Hindi ibebenta o gagamitin sa marketing ang inyong impormasyon. Alinsunod sa umiiral na ' +
        'batas, maaari kayong humiling na makita, iwasto, o ipabura ang inyong impormasyon, o ' +
        'bawiin ang pahintulot, sa pamamagitan ng pakikipag-ugnayan sa FloodSense research team ' +
        'sa survey administrator.'
    );

  const consentItem = form
    .addMultipleChoiceItem()
    .setTitle('Consent / Pahintulot')
    .setHelpText(
      'If you do not agree, the form will be submitted and the survey will end. / ' +
        'Kung hindi kayo pumapayag, ipapasa ang form at matatapos ang survey.'
    )
    .setRequired(true);

  const demographicsPage = form
    .addPageBreakItem()
    .setTitle('Demographic Information / Impormasyong Demograpiko')
    .setHelpText(
      'Please provide the following information. You may enter Prefer not to answer for any ' +
        'item you do not wish to disclose.\n\n' +
        'Ibigay ang sumusunod na impormasyon. Maaari ninyong ilagay ang Mas nais kong hindi ' +
        'sumagot sa anumang aytem na ayaw ninyong ibahagi.'
    );

  consentItem.setChoices([
    consentItem.createChoice(
      'I have read and understood the notice, and I voluntarily agree to participate. / ' +
        'Nabasa at naunawaan ko ang pahayag, at kusang-loob akong pumapayag na lumahok.',
      demographicsPage
    ),
    consentItem.createChoice(
      'I do not agree to participate. / Hindi ako pumapayag na lumahok.',
      FormApp.PageNavigationType.SUBMIT
    ),
  ]);

  addTextQuestion_(
    form,
    'D1. Full name / Buong pangalan',
    'Enter Prefer not to answer if you do not wish to disclose your name. / ' +
      'Ilagay ang Mas nais kong hindi sumagot kung ayaw ninyong ibahagi ang inyong pangalan.',
    true
  );

  addTextQuestion_(
    form,
    'D2. Age / Edad',
    'Enter your age in years, or enter Prefer not to answer. / ' +
      'Ilagay ang inyong edad sa taon, o ilagay ang Mas nais kong hindi sumagot.',
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'D3. Sex or gender / Kasarian',
    '',
    [
      'Male / Lalaki',
      'Female / Babae',
      'Prefer to self-describe / Nais ilarawan ang sarili',
      'Prefer not to answer / Mas nais kong hindi sumagot',
    ],
    true
  );

  addTextQuestion_(
    form,
    'D3a. If you selected Prefer to self-describe, please specify. / ' +
      'Kung pinili ninyo ang Nais ilarawan ang sarili, pakisulat dito.',
    '',
    false
  );

  addTextQuestion_(
    form,
    'D4. Barangay or community and city / Barangay o komunidad at lungsod',
    'Provide only your general location; do not enter a house number or exact street address. / ' +
      'Pangkalahatang lugar lamang ang ilagay; huwag ilagay ang numero ng bahay o eksaktong address.',
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'D5. How long have you lived in your current barangay or community?\n' +
      'Gaano na kayo katagal nakatira sa inyong kasalukuyang barangay o komunidad?',
    '',
    [
      'Less than 1 year / Mas mababa sa 1 taon',
      '1 to 5 years / 1 hanggang 5 taon',
      '6 to 10 years / 6 hanggang 10 taon',
      'More than 10 years / Higit sa 10 taon',
      'Prefer not to answer / Mas nais kong hindi sumagot',
    ],
    true
  );

  const q1Page = form
    .addPageBreakItem()
    .setTitle('Part I — Eligibility and Flood Experience / Pagiging Kwalipikado at Karanasan sa Baha');

  const q1 = form
    .addMultipleChoiceItem()
    .setTitle('Q1. Are you 18 years old or older?\nKayo ba ay 18 taong gulang o mas matanda?')
    .setHelpText(
      'If No or Prefer not to answer is selected, the survey will end. / ' +
        'Kung Hindi o Mas nais kong hindi sumagot ang napili, matatapos ang survey.'
    )
    .setRequired(true);

  const q2Page = form
    .addPageBreakItem()
    .setTitle('Eligibility Check / Pagsusuri ng Pagiging Kwalipikado');

  q1.setChoices([
    q1.createChoice('Yes / Oo', q2Page),
    q1.createChoice('No / Hindi', FormApp.PageNavigationType.SUBMIT),
    q1.createChoice(
      'Prefer not to answer / Mas nais kong hindi sumagot',
      FormApp.PageNavigationType.SUBMIT
    ),
  ]);

  const q2 = form
    .addMultipleChoiceItem()
    .setTitle(
      'Q2. Do you currently live in a Bacoor City community included in this study?\n' +
        'Kasalukuyan ba kayong nakatira sa isang komunidad sa Bacoor City na kasama sa pag-aaral na ito?'
    )
    .setHelpText(
      'If No is selected, the survey will end. / Kung Hindi ang napili, matatapos ang survey.'
    )
    .setRequired(true);

  const q3Page = form
    .addPageBreakItem()
    .setTitle('Flood Experience / Karanasan sa Baha');

  q2.setChoices([
    q2.createChoice('Yes / Oo', q3Page),
    q2.createChoice('No / Hindi', FormApp.PageNavigationType.SUBMIT),
    q2.createChoice('I am not sure / Hindi ako tiyak', q3Page),
  ]);

  addMultipleChoiceQuestion_(
    form,
    'Q3. During the past three years, how often has flooding or heavy rain directly affected ' +
      'your household, travel, work, or studies?\n' +
      'Sa nakalipas na tatlong taon, gaano kadalas direktang naapektuhan ng baha o malakas na ' +
      'ulan ang inyong sambahayan, pagbiyahe, trabaho, o pag-aaral?',
    '',
    [
      'Very often / Napakadalas',
      'Often / Madalas',
      'Sometimes / Paminsan-minsan',
      'Rarely / Bihira',
      'Never / Hindi kailanman',
      'I cannot recall / Hindi ko maalala',
    ],
    true
  );

  form
    .addPageBreakItem()
    .setTitle('Part II — Current Flood-Information Experience / Kasalukuyang Karanasan sa Impormasyon Tungkol sa Baha');

  addMultipleChoiceQuestion_(
    form,
    'Q4. Which source do you usually check first for flood-related information?\n' +
      'Aling pinagmumulan ang karaniwan ninyong unang tinitingnan para sa impormasyon tungkol sa baha?',
    '',
    [
      'Official government social-media page / Opisyal na social-media page ng pamahalaan',
      'Barangay announcement or public-address system / Anunsiyo o public-address system ng barangay',
      'Text or SMS alert / Text o SMS alert',
      'Television or radio / Telebisyon o radyo',
      'Government or weather website / Website ng pamahalaan o serbisyong pangpanahon',
      'Family, neighbors, or community group / Pamilya, kapitbahay, o grupong pangkomunidad',
      'Personal observation / Sariling pagmamasid',
      'I do not usually check any source / Wala akong karaniwang tinitingnang pinagmulan',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q5. During flooding or heavy rain, how often are you able to obtain the information you need?\n' +
      'Kapag may baha o malakas na ulan, gaano kadalas ninyong nakukuha ang impormasyong kailangan ninyo?',
    '',
    [
      'Always / Palagi',
      'Often / Madalas',
      'Sometimes / Paminsan-minsan',
      'Rarely / Bihira',
      'Never / Hindi kailanman',
      'I usually do not need flood information / Karaniwan ay hindi ko kailangan ang impormasyon',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q6. When does useful flood information usually reach you?\n' +
      'Kailan karaniwang dumarating sa inyo ang kapaki-pakinabang na impormasyon tungkol sa baha?',
    '',
    [
      'Early enough to prepare or decide / Sapat ang aga upang makapaghanda o makapagdesisyon',
      'Shortly before I need to act / Kaunting oras bago ko kailangang kumilos',
      'After I have already started acting / Pagkatapos kong magsimulang kumilos',
      'After flooding has already affected me / Pagkatapos akong maapektuhan ng baha',
      'I usually receive no useful information / Karaniwan ay wala akong natatanggap na kapaki-pakinabang na impormasyon',
      'I cannot recall / Hindi ko maalala',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q7. How easy or difficult is the flood information you receive to understand?\n' +
      'Gaano kadali o kahirap unawain ang impormasyong natatanggap ninyo tungkol sa baha?',
    '',
    [
      'Very easy / Napakadali',
      'Easy / Madali',
      'Neither easy nor difficult / Hindi madali at hindi rin mahirap',
      'Difficult / Mahirap',
      'Very difficult / Napakahirap',
      'I do not receive flood information / Wala akong natatanggap na impormasyon',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q8. How well does the information identify the location that may be affected?\n' +
      'Gaano kahusay tinutukoy ng impormasyon ang lugar na maaaring maapektuhan?',
    '',
    [
      'It gives all the location details I need / Ibinibigay nito ang lahat ng detalyeng kailangan ko',
      'It gives some, but not all, location details / May ibinibigay itong detalye ngunit hindi kumpleto',
      'It is too general to help me identify the affected location / Masyado itong pangkalahatan upang matukoy ang apektadong lugar',
      'Location details are not important for my decisions / Hindi mahalaga sa aking desisyon ang detalye ng lugar',
      'I do not receive flood information / Wala akong natatanggap na impormasyon',
      'I am not sure / Hindi ako tiyak',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q9. How often do different flood-information sources provide conflicting or inconsistent information?\n' +
      'Gaano kadalas nagbibigay ng magkakasalungat o hindi magkakatugmang impormasyon ang iba’t ibang pinagmumulan?',
    '',
    [
      'Very often / Napakadalas',
      'Often / Madalas',
      'Sometimes / Paminsan-minsan',
      'Rarely / Bihira',
      'Never / Hindi kailanman',
      'I use only one source / Isang pinagmumulan lamang ang ginagamit ko',
      'I am not sure / Hindi ako tiyak',
    ],
    true
  );

  addMultipleChoiceQuestion_(
    form,
    'Q10. Which statement best describes how the available information affects your decisions?\n' +
      'Aling pahayag ang pinakamahusay na naglalarawan kung paano nakaaapekto sa inyong mga ' +
      'desisyon ang impormasyong mayroon kayo?',
    '',
    [
      'It usually helps me decide in time / Karaniwan ay nakatutulong ito upang makapagdesisyon ako sa tamang oras',
      'It helps, but sometimes arrives too late / Nakatutulong ito ngunit kung minsan ay huli nang dumarating',
      'Information is available but is not useful for my decision / May impormasyon ngunit hindi ito kapaki-pakinabang sa aking desisyon',
      'I usually cannot obtain the information I need / Karaniwan ay hindi ko nakukuha ang impormasyong kailangan ko',
      'I usually do not need information to decide / Karaniwan ay hindi ko kailangan ng impormasyon upang makapagdesisyon',
      'I am not sure / Hindi ako tiyak',
    ],
    true
  );

  addCheckboxQuestion_(
    form,
    'Q11. What problems do you experience with flood information?\n' +
      'Anong mga problema ang nararanasan ninyo sa impormasyon tungkol sa baha?',
    'Select all that apply. / Pumili ng lahat na naaangkop.',
    [
      'It arrives too late / Huli itong dumarating',
      'It is difficult to understand / Mahirap itong unawain',
      'It does not identify the affected location clearly / Hindi nito malinaw na tinutukoy ang apektadong lugar',
      "Different sources provide inconsistent information / Hindi magkakatugma ang impormasyon mula sa iba't ibang pinagmumulan",
      'I cannot access it when needed / Hindi ko ito ma-access kapag kailangan',
      'It appears outdated / Tila luma na ang impormasyon',
      'I do not usually experience a problem / Karaniwan ay wala akong nararanasang problema',
      'I am not sure / Hindi ako tiyak',
    ],
    true
  );

  addCheckboxQuestion_(
    form,
    'Q12. For planning before heavy rain, how far in advance is flood-related information useful to you?\n' +
      'Para sa paghahanda bago ang malakas na ulan, gaano kaagang impormasyon tungkol sa baha ' +
      'ang kapaki-pakinabang sa inyo?',
    'Select all that apply. / Pumili ng lahat na naaangkop.',
    [
      'Less than 30 minutes before / Wala pang 30 minuto bago',
      'About 30 minutes to 1 hour before / Humigit-kumulang 30 minuto hanggang 1 oras bago',
      'About 2 to 3 hours before / Humigit-kumulang 2 hanggang 3 oras bago',
      'About 4 to 6 hours before / Humigit-kumulang 4 hanggang 6 na oras bago',
      'More than 6 hours before / Mahigit 6 na oras bago',
      'It depends on the situation / Depende sa sitwasyon',
      'I do not need advance information / Hindi ko kailangan ng maagang impormasyon',
    ],
    true
  );

  const q13 = addCheckboxQuestion_(
    form,
    'Q13. Which of these flood information types help you prepare?\n' +
      'Alin sa mga sumusunod na uri ng impormasyon tungkol sa baha ang tumutulong sa inyong paghahanda?',
    'Select 1 or 2 answers. / Pumili ng 1 o 2 sagot.',
    [
      'Advance flood warning or advisory / Maagang babala o abiso tungkol sa baha',
      'Whether my area may be susceptible to flooding / Kung maaaring bahain ang aking lugar',
      'Expected road or route conditions before flooding occurs / Inaasahang kalagayan ng kalsada o ruta bago mangyari ang baha',
      'Expected water depth in flood-prone areas / Inaasahang lalim ng tubig sa mga lugar na madalas bahain',
      'Recommended or planned evacuation routes / Inirerekomenda o pinlanong mga rutang pang-evacuation',
      'Verified evacuation-center information / Beripikadong impormasyon tungkol sa evacuation center',
      'Household preparedness instructions / Mga tagubilin sa paghahanda ng sambahayan',
      'Emergency contact information / Impormasyon sa emergency contact',
      'I do not need additional information / Hindi ko kailangan ng dagdag na impormasyon',
    ],
    true
  );

  q13.setValidation(
    FormApp.createCheckboxValidation()
      .requireSelectAtLeast(1)
      .requireSelectAtMost(2)
      .setHelpText('Select no more than two answers. / Pumili ng hindi hihigit sa dalawang sagot.')
      .build()
  );

  addMultipleChoiceQuestion_(
    form,
    'Q14. Overall, how well do your current sources meet your flood-information needs?\n' +
      'Sa kabuuan, gaano kahusay natutugunan ng inyong kasalukuyang mga pinagmumulan ang ' +
      'pangangailangan ninyo sa impormasyon tungkol sa baha?',
    '',
    [
      'Completely / Ganap',
      'Mostly / Karamihan',
      'Partly / Bahagya',
      'Very little / Kaunti lamang',
      'Not at all / Hindi talaga',
      'I do not need flood information / Hindi ko kailangan ng impormasyon tungkol sa baha',
    ],
    true
  );

  form
    .addPageBreakItem()
    .setTitle('Part III — Technology Access and Preference / Access at Kagustuhan sa Teknolohiya');

  addMultipleChoiceQuestion_(
    form,
    'Q15. Which statement best describes your usual access to a mobile phone?\n' +
      'Aling pahayag ang pinakamahusay na naglalarawan sa karaniwan ninyong access sa mobile phone?',
    '',
    [
      'I regularly use my own smartphone / Regular kong ginagamit ang sarili kong smartphone',
      'I regularly use a shared smartphone / Regular kong ginagamit ang smartphone na pinagsasaluhan',
      'I use a basic phone without applications / Basic phone na walang applications ang ginagamit ko',
      'I do not regularly have access to a mobile phone / Wala akong regular na access sa mobile phone',
      'Prefer not to answer / Mas nais kong hindi sumagot',
    ],
    true
  );

  addCheckboxQuestion_(
    form,
    'Q16. Which of the following methods helps you receive flood information?\n' +
      'Alin sa mga sumusunod na paraan ang tumutulong sa inyo upang matanggap ang impormasyon tungkol sa baha?',
    'Select all that apply. / Pumili ng lahat na naaangkop.',
    [
      'Text or SMS / Text o SMS',
      'Official social media / Opisyal na social media',
      'Barangay public-address announcement / Public-address announcement ng barangay',
      'Television or radio / Telebisyon o radyo',
      'Mobile application / Mobile application',
      'Government website / Website ng pamahalaan',
      'Telephone hotline / Telephone hotline',
      'Family or community network / Pamilya o network ng komunidad',
      'None of these methods help me / Wala sa mga ito ang nakatutulong sa akin',
      'I do not want to receive additional information / Ayaw ko ng dagdag na impormasyon',
    ],
    true
  );

  const editUrl = form.getEditUrl();
  const respondentUrl = form.getPublishedUrl();

  Logger.log('Google Form created successfully.');
  Logger.log('Editor URL: ' + editUrl);
  Logger.log('Respondent URL: ' + respondentUrl);

  return {
    formId: form.getId(),
    editUrl: editUrl,
    respondentUrl: respondentUrl,
  };
}


function addTextQuestion_(form, title, helpText, required) {
  const item = form.addTextItem().setTitle(title).setRequired(required);
  if (helpText) {
    item.setHelpText(helpText);
  }
  return item;
}


function addMultipleChoiceQuestion_(form, title, helpText, choices, required) {
  const item = form
    .addMultipleChoiceItem()
    .setTitle(title)
    .setChoiceValues(choices)
    .setRequired(required);
  if (helpText) {
    item.setHelpText(helpText);
  }
  return item;
}


function addCheckboxQuestion_(form, title, helpText, choices, required) {
  const item = form
    .addCheckboxItem()
    .setTitle(title)
    .setChoiceValues(choices)
    .setRequired(required);
  if (helpText) {
    item.setHelpText(helpText);
  }
  return item;
}
