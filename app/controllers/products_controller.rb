#require 'powerpoint'
class ProductsController < ApplicationController
  before_action :set_product, only: [:show, :edit, :update, :destroy]

  # GET /products
  # GET /products.json
  def index
    #@deck = Powerpoint::Presentation.new
    # Creating an introduction slide:
    # title = 'Bicycle Of the Mind'
    # subtitle = 'created by Steve Jobs'
    # @deck.add_intro title, subtitle

    # Creating a text-only slide:
    # Title must be a string.
    # Content must be an array of strings that will be displayed as bullet items.
    # title = 'Why Mac?'
    # content = ['Its cool!', 'Its light.']
    # @deck.add_textual_slide title, content

    # Creating an image Slide:
    # It will contain a title as string.
    # and an embeded image
    #title = 'Everyone loves Macs:'
    #subtitle = 'created by Steve Jobs'
    #content = ['Its cool!', 'Its light.']
    #image_path = ActionController::Base.helpers.asset_path('/app/assets/images/ss.png').to_s
    #image = view_context.image_path 'ss.png'
    #url = 'http://localhost:3000' + image
    #image = view_context.image_path 'https://res.cloudinary.com/indoexchanger/image/upload/v1501168483/qxbvro2yhvibid0ra5rp.jpg'
    #puts image
    #@deck.add_pictorial_slide title, url
    #@deck.add_textual_slide title, subtitle
    #

    # Specifying coordinates and image size for an embeded image.
    # x and y values define the position of the image on the slide.
    # cx and cy define the width and height of the image.
    # x, y, cx, cy are in points. Each pixel is 12700 points.
    # coordinates parameter is optional.
    # coords = {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
    # @deck.add_pictorial_slide title, image_path, coords

    # Saving the pptx file to the current directory.
    #@deck.save('mps.pptx')
    # @products = Product.all
    # 
    
  @presentation =  RubySlides::Presentation.new
      
  chart_title = "Chart Slide"
  chart_series = [
    {
      column: "Col1",
      rows:   ["Lorem", "Ipsum", "Dolar", "Ismet"],
      values: ["1", "3", "5", "7"]
    },
    {
      column: "Col2",
      rows:   ["A", "B", "C", "D"],
      values: ["2", "4", "6", "8"]
    }
  ]
  @presentation.chart_slide chart_title, chart_series

  @presentation.save('mps.pptx')

    @products = Product.order('created_at DESC')
    respond_to do |format|
      format.html
      format.xlsx {
        response.headers['Content-Disposition'] = 'attachment; filename="all_products.xlsx"'
      }
    end
  end

  # GET /products/1
  # GET /products/1.json
  def show
  end

  # GET /products/new
  def new
    @product = Product.new
  end

  # GET /products/1/edit
  def edit
  end

  # POST /products
  # POST /products.json
  def create
    @product = Product.new(product_params)

    respond_to do |format|
      if @product.save
        format.html { redirect_to @product, notice: 'Product was successfully created.' }
        format.json { render :show, status: :created, location: @product }
      else
        format.html { render :new }
        format.json { render json: @product.errors, status: :unprocessable_entity }
      end
    end
  end

  # PATCH/PUT /products/1
  # PATCH/PUT /products/1.json
  def update
    respond_to do |format|
      if @product.update(product_params)
        format.html { redirect_to @product, notice: 'Product was successfully updated.' }
        format.json { render :show, status: :ok, location: @product }
      else
        format.html { render :edit }
        format.json { render json: @product.errors, status: :unprocessable_entity }
      end
    end
  end

  # DELETE /products/1
  # DELETE /products/1.json
  def destroy
    @product.destroy
    respond_to do |format|
      format.html { redirect_to products_url, notice: 'Product was successfully destroyed.' }
      format.json { head :no_content }
    end
  end

  private
    # Use callbacks to share common setup or constraints between actions.
    def set_product
      @product = Product.find(params[:id])
    end

    # Never trust parameters from the scary internet, only allow the white list through.
    def product_params
      params.require(:product).permit(:title, :price)
    end
end
