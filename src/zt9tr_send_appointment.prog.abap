REPORT zt9tr_send_appointment.

INCLUDE <cntn01>.
TYPE-POOLS: sccon.

PARAMETERS p_orga  TYPE xubname    DEFAULT sy-uname OBLIGATORY.
PARAMETERS p_mail  TYPE ad_smtpadr DEFAULT 'ewf@inwerken.de' OBLIGATORY.
PARAMETERS p_mail2 TYPE ad_smtpadr DEFAULT 'lmr@inwerken.de'.
PARAMETERS p_title TYPE sc_txtshor DEFAULT 'Geschäftsessen'.
PARAMETERS p_loc   TYPE sc_room    DEFAULT 'La Civetta'.
PARAMETERS p_date  TYPE sy-datum   DEFAULT sy-datum.
PARAMETERS p_from  TYPE sc_timefro DEFAULT '120000'.
PARAMETERS p_to    TYPE sc_timeto  DEFAULT '130000'.
SELECTION-SCREEN BEGIN OF BLOCK body WITH FRAME TITLE TEXT-bdy.
PARAMETERS p_line1 TYPE so_text255 DEFAULT 'Wichtiges Essen'.
PARAMETERS p_line2 TYPE so_text255 DEFAULT 'Schickes Hemd anziehen'.
PARAMETERS p_line3 TYPE so_text255 DEFAULT 'Blumen mitbringen'.
SELECTION-SCREEN END OF BLOCK body.

CLASS main DEFINITION.
  PUBLIC SECTION.

    CONSTANTS c_status_confirmation_never    TYPE bcs_rqst VALUE 'N'.  "Never
    CONSTANTS c_status_confirmation_on_error TYPE bcs_rqst VALUE 'E'.  "Only if errors occur
    CONSTANTS c_status_confirmation_if_sent  TYPE bcs_rqst VALUE 'D'.  "If sent
    CONSTANTS c_status_confirmation_if_read  TYPE bcs_rqst VALUE 'R'.  "If read
    CONSTANTS c_status_confirmation_always   TYPE bcs_rqst VALUE 'A'.  "Always

    METHODS start.
  PRIVATE SECTION.
    DATA appointment       TYPE REF TO cl_appointment.
    DATA participant       TYPE scspart.

    METHODS add_participant IMPORTING i_mail_address TYPE clike.
ENDCLASS.

CLASS main IMPLEMENTATION.
  METHOD start.
    appointment = NEW #( ).

    "MEETING, VACATION, CUSTOMER, ABSENT
    appointment->set_type( 'ZREMINDER' ).

    appointment->set_organizer( organizer = p_orga ).

    add_participant( p_mail ).
    IF p_mail2 IS NOT INITIAL.
      add_participant( p_mail2 ).
    ENDIF.

    " add detail body text
    appointment->set_text( VALUE #(
     ( line = p_line1 ) ( line = cl_abap_char_utilities=>cr_lf )
     ( line = p_line2 ) ( line = cl_abap_char_utilities=>cr_lf )
     ( line = p_line3 )
     ) ).

    " set title and location
    appointment->set_title( p_title ).
    appointment->set_location( p_loc ).

    " set date and time using default settings
    " date_to will be the same as date_from
    " time zone will be the one from the user master records settings
    appointment->set_date( date_from = p_date
                           time_from = p_from
                           time_to   = p_to ). "Central european time
    " set it to a high priority meeting
    appointment->set_priority( sccon_prio_very_high ).

    " this meeting is not yet confirmed
    appointment->set_status( sccon_status_planned ).
    " Important to set this one to space. Otherwise SAP will send a not user-friendly e-mail
    appointment->save( send_invitation = space ).

    TRY.
        " Now that we have the appointment, we can send a good one for outlook by switching to BCS
        DATA(send_request) = appointment->create_send_request( ).
        DATA(recipient) = cl_cam_address_bcs=>create_internet_address( p_mail ).
        send_request->add_recipient( i_recipient = recipient i_copy = abap_true ).
        IF p_mail2 IS NOT INITIAL.
          DATA(recipient2) = cl_cam_address_bcs=>create_internet_address( p_mail2 ).
          send_request->add_recipient( i_recipient = recipient2 i_copy = abap_true ).
        ENDIF.
      CATCH cx_address_bcs INTO DATA(error_address).
        MESSAGE error_address TYPE 'I'.
        RETURN.
      CATCH cx_send_req_bcs INTO DATA(error_add_recipient).
        MESSAGE error_add_recipient TYPE 'I'.
        RETURN.
      CATCH cx_bcs INTO DATA(error_create_send_request).
        MESSAGE error_create_send_request TYPE 'I'.
        RETURN.
    ENDTRY.

    TRY.
        " don't request read/delivery receipts
        send_request->set_status_attributes(
          i_requested_status = c_status_confirmation_never
          i_status_mail      = c_status_confirmation_never ).
        "sent mail immediately
        send_request->set_send_immediately( abap_true ).
        " Send it to the world
        DATA(appointment_sent) = send_request->send( i_with_error_screen = abap_true ).
        IF appointment_sent = abap_true.
          COMMIT WORK AND WAIT.
          MESSAGE 'Einladung verschickt' TYPE 'S'.
        ELSE.
          MESSAGE 'Fehler beim Senden der Einladung' TYPE 'I'.
        ENDIF.
      CATCH cx_send_req_bcs INTO DATA(error_send).
        MESSAGE error_send TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD add_participant.

    DATA address           TYPE swc_object.
    DATA address_container TYPE STANDARD TABLE OF swcont.

    "set an internet address as a second partcipant of that appointment
    swc_create_object address 'ADDRESS' space.
    swc_set_element address_container 'AddressString'  i_mail_address.
    swc_set_element address_container 'TypeId' 'U'.
    swc_set_element address_container 'NoAdradmi' 'X'.
    swc_set_element address_container 'NoIntern' 'X'.
    swc_call_method address 'Create' address_container.
    CHECK sy-subrc = 0.
    "get key and type of object
    swc_get_object_key address participant-objkey.
    CHECK sy-subrc = 0.
    swc_get_object_type address participant-objtype.
    CHECK sy-subrc = 0.
    participant-send_mail = abap_true.
    appointment->add_participant( participant = participant ).

  ENDMETHOD.
ENDCLASS.

START-OF-SELECTION.

  NEW main( )->start( ).
