#include <opencv2/opencv.hpp>
#include <cairo/cairo.h>
#include <iostream>
#include <sstream>
#include <iomanip>

int main() {
    int n;
    std::cout << "Enter the size of the multiplication table: ";
    std::cin >> n;

    int cellSize = 80; // pixel size per cell
    int width = cellSize * (n + 1);
    int height = cellSize * (n + 1);

    // -------------------- PNG using OpenCV --------------------
    cv::Mat img(height, width, CV_8UC3, cv::Scalar(255,255,255)); // white background
    cv::Scalar textColor(0,0,0); // black

    // Draw grid
    for(int i=0;i<=n;i++){
        cv::line(img, cv::Point(cellSize*i,0), cv::Point(cellSize*i,height), cv::Scalar(0,0,0), 2);
        cv::line(img, cv::Point(0, cellSize*i), cv::Point(width, cellSize*i), cv::Scalar(0,0,0), 2);
    }

    // Fill numbers
    for(int i=0;i<=n;i++){
        for(int j=0;j<=n;j++){
            std::stringstream ss;
            if(i==0 && j==0) ss << "*";
            else if(i==0) ss << j;
            else if(j==0) ss << i;
            else ss << i*j;

            std::string text = ss.str();
            int baseline=0;
            cv::Size textSize = cv::getTextSize(text, cv::FONT_HERSHEY_SIMPLEX, 1, 2, &baseline);
            int x = j*cellSize + (cellSize - textSize.width)/2;
            int y = i*cellSize + (cellSize + textSize.height)/2;

            cv::putText(img, text, cv::Point(x,y), cv::FONT_HERSHEY_SIMPLEX, 1, textColor, 2);
        }
    }

    cv::imwrite("multiplication_table.png", img);
    std::cout << "PNG file saved as multiplication_table.png\n";

    // -------------------- PDF using Cairo --------------------
    cairo_surface_t *surface = cairo_pdf_surface_create("multiplication_table.pdf", width, height);
    cairo_t *cr = cairo_create(surface);

    // White background
    cairo_set_source_rgb(cr, 1,1,1);
    cairo_paint(cr);

    // Draw grid
    cairo_set_source_rgb(cr, 0,0,0);
    cairo_set_line_width(cr, 1.5);
    for(int i=0;i<=n;i++){
        cairo_move_to(cr, 0, i*cellSize);
        cairo_line_to(cr, width, i*cellSize);
        cairo_move_to(cr, i*cellSize, 0);
        cairo_line_to(cr, i*cellSize, height);
    }
    cairo_stroke(cr);

    // Draw text
    cairo_select_font_face(cr, "Sans", CAIRO_FONT_SLANT_NORMAL, CAIRO_FONT_WEIGHT_BOLD);
    cairo_set_font_size(cr, 20);

    for(int i=0;i<=n;i++){
        for(int j=0;j<=n;j++){
            std::stringstream ss;
            if(i==0 && j==0) ss << "*";
            else if(i==0) ss << j;
            else if(j==0) ss << i;
            else ss << i*j;
            std::string text = ss.str();

            cairo_text_extents_t extents;
            cairo_text_extents(cr, text.c_str(), &extents);

            double x = j*cellSize + (cellSize - extents.width)/2 - extents.x_bearing;
            double y = i*cellSize + (cellSize - extents.height)/2 - extents.y_bearing;
            cairo_move_to(cr, x, y);
            cairo_show_text(cr, text.c_str());
        }
    }

    cairo_destroy(cr);
    cairo_surface_destroy(surface);

    std::cout << "PDF file saved as multiplication_table.pdf\n";

    return 0;
}
