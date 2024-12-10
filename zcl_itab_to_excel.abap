CLASS zcl_send_oo_email DEFINITION
  PUBLIC FINAL CREATE PUBLIC .

  PUBLIC SECTION.
    CLASS-METHODS send_email .
ENDCLASS.

CLASS zcl_send_oo_email IMPLEMENTATION.

  METHOD send_email.

    "Get Data
    SELECT * FROM /dmo/flight INTO TABLE @DATA(lt_data).
    GET REFERENCE OF lt_data INTO DATA(lo_data_ref).
    DATA(lv_xstring) = NEW zcl_itab_to_excel( )->itab_to_xstring( lo_data_ref ).

*--- Email code starts here
    TRY.
        "Create send request
        DATA(lo_send_request) = cl_bcs=>create_persistent( ).

        "Create mail body
        DATA(lt_body) = VALUE bcsy_text(
                          ( line = 'Dear Recipient,' ) ( )
                          ( line = 'PFA flight details file.' ) ( )
                          ( line = 'Thank You' )
                        ).

        "Set up document object
        DATA(lo_document) = cl_document_bcs=>create_document(
                              i_type = 'RAW'
                              i_text = lt_body
                              i_subject = 'Flight Details' ).

        "Add attachment
        lo_document->add_attachment(
            i_attachment_type    = 'xls'
            i_attachment_size    = CONV #( xstrlen( lv_xstring ) )
            i_attachment_subject = 'Flight Details'
            i_attachment_header  = VALUE #( ( line = 'Flights.xlsx' ) )
            i_att_content_hex    = cl_bcs_convert=>xstring_to_solix( lv_xstring )
         ).

        "Add document to send request
        lo_send_request->set_document( lo_document ).

        "Set sender
        lo_send_request->set_sender(
          cl_cam_address_bcs=>create_internet_address(
            i_address_string = CONV #( 'sender@dummy.com' )
          )
        ).

        "Set Recipient | This method has options to set CC/BCC as well
        lo_send_request->add_recipient(
          i_recipient = cl_cam_address_bcs=>create_internet_address(
                          i_address_string = CONV #( 'recipient@dummy.com' )
                        )
          i_express   = abap_true ).

        "Send Email
        DATA(lv_sent_to_all) = lo_send_request->send( ).
        COMMIT WORK.

      CATCH cx_send_req_bcs INTO DATA(lx_req_bsc).
        "Error handling
      CATCH cx_document_bcs INTO DATA(lx_doc_bcs).
        "Error handling
      CATCH cx_address_bcs  INTO DATA(lx_add_bcs).
        "Error handling
    ENDTRY.

  ENDMETHOD.
ENDCLASS.
